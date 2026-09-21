<#
.SYNOPSIS
  New vCenter side (step 4): register VMs / templates back into the inventory and
  place them directly into their original folders. Two modes:

  Mode B (main path) -UnregisteredCsv
    Register exactly the VMs listed in unregistered.csv produced by
    Unregister-VmFromOldVc. vmx path, name, template flag and original folder all
    come from the CSV; the CSV can be edited in Excel before running.

  Mode A -Datastore
    Scan the given datastores for every .vmx / .vmtx that is not yet in the
    inventory and register them (orphan clean-up, or when no CSV exists).
    Use -PlacementCsv (vm-placement.csv from Export-VcMeta) to place them into folders.

.DESCRIPTION
  Reconciliation at the end: VMs in the CSV (or manifest) that are still not on
  the new vCenter are listed as PendingOnSource, so batches can be tracked.
  Already registered vmx paths are skipped, so the script can be re-run safely.
  .vmtx files are registered as templates. Nothing is powered on.

  Optional: -MetaDir <export dir> also applies attribute values / notes / tag
  assignments to the registered VMs in the same run (step 5 merged in).

  Self-signed vCenter certificates are accepted automatically. The PowerCLI CEIP
  prompt is suppressed so the script never blocks on first run.

.EXAMPLE
  # Main path, dry run first
  .\Register-VmxFromDatastore.ps1 -Server vc-new.example.com -Password '<pw>' -Cluster cl01 -UnregisteredCsv .\export-A\unregistered.csv -DryRun

.EXAMPLE
  # Main path, for real (register + place into original folders)
  .\Register-VmxFromDatastore.ps1 -Server vc-new.example.com -Password '<pw>' -Cluster cl01 -UnregisteredCsv .\export-A\unregistered.csv

.EXAMPLE
  # Scan a datastore (no CSV available)
  .\Register-VmxFromDatastore.ps1 -Server vc-new.example.com -Password '<pw>' -Datastore ds01 -Cluster cl01 -PlacementCsv .\export-A\vm-placement.csv -DryRun

.EXAMPLE
  # Optional: merge step 5 (attribute values / notes / tags) into this run
  .\Register-VmxFromDatastore.ps1 ... -UnregisteredCsv .\export-A\unregistered.csv -MetaDir .\export-A -DatacenterMap 'DC-A=DC-B'
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$Server,
    [string]$User,
    [string]$Password,
    [pscredential]$Credential,

    [string[]]$Datastore,              # mode A: scan these datastores for vmx/vmtx
    [string]$UnregisteredCsv,          # mode B (main path): unregistered.csv from Unregister-VmFromOldVc
    [string]$Cluster,                  # register into this cluster (hosts that mount the datastore, round-robin)
    [string]$VMHost,                   # or a single host
    [string]$ResourcePool,             # default: root resource pool of the host's cluster / host

    [string]$Folder,                   # mode A default target VM folder (path relative to the datacenter, e.g. 'Linux'); default = DC root
    [string]$PlacementCsv,             # mode A: vm-placement.csv from Export-VcMeta, moves each VM into its folder after registration
    [switch]$CreateFolders,            # create folders that do not exist

    # --- Optional: apply folders / attributes / notes / tags from the old vCenter in the same run ---
    [string]$SourceServer,             # fetch live from the old vCenter (runs Export-VcMeta into -MetaDir); VMs must still be in its inventory
    [string]$SourceUser = 'administrator@vsphere.local',
    [string]$SourcePassword,
    [string]$MetaDir,                  # or an existing export directory (-OutDir of Export-VcMeta) -- recommended: export first, then unregister
    [string[]]$DatacenterMap,          # old=new datacenter name mapping, e.g. 'DC-A=DC-B'
    [string[]]$SourceDatacenter,       # only these datacenters from the old vCenter

    [string[]]$Include,                # only names matching these wildcards, e.g. 'web*','db01'
    [string[]]$Exclude = @('vCLS*'),   # exclude (vCLS skipped by default)
    [switch]$NameFromFile,             # use the vmx file name as VM name (default: displayName from the vmx, file name as fallback)
    [switch]$NoTemplates,
    [switch]$DryRun,
    [string]$ReportPath
)

$ErrorActionPreference = 'Stop'
if (-not $Datastore -and -not $UnregisteredCsv) { throw "Give either -UnregisteredCsv (main path) or -Datastore (scan mode)" }
$enc = if ($PSVersionTable.PSVersion.Major -ge 6) { 'utf8BOM' } else { 'UTF8' }
if (-not $ReportPath) { $ReportPath = Join-Path (Get-Location) ('register-report-{0:yyyyMMdd-HHmmss}.csv' -f (Get-Date)) }

$script:Report = New-Object System.Collections.ArrayList
$script:Stats  = @{}
function Add-Result {
    param([string]$Action,[string]$Target,[string]$Detail = '')
    [void]$script:Report.Add([pscustomobject]@{ Action = $Action; Target = $Target; Detail = $Detail })
    if (-not $script:Stats.ContainsKey($Action)) { $script:Stats[$Action] = 0 }
    $script:Stats[$Action]++
    $tag = switch ($Action) { 'Registered' { '[+]' } 'WouldRegister' { '[~]' } 'Failed' { '[X]' } default { '[ ]' } }
    Write-Host ("  {0} {1,-18} {2}  {3}" -f $tag, $Action, $Target, $Detail)
}
function Test-NameMatch {
    param([string]$Name)
    if ($Include) { $hit = $false; foreach ($p in $Include) { if ($Name -like $p) { $hit = $true; break } }; if (-not $hit) { return $false } }
    foreach ($p in $Exclude) { if ($Name -like $p) { return $false } }
    return $true
}

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$script:RegisteredNames = New-Object System.Collections.ArrayList

# --- PowerCLI: accept self-signed certificates, suppress CEIP prompt ---
try { Import-Module VMware.VimAutomation.Core -ErrorAction Stop } catch { throw "VMware.PowerCLI is not installed. Run: Install-Module VMware.PowerCLI -Scope CurrentUser" }
try { Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -ParticipateInCeip $false -DisplayDeprecationWarnings $false -Scope User -Confirm:$false | Out-Null } catch { }
Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -DisplayDeprecationWarnings $false -Scope Session -Confirm:$false | Out-Null

# --- Optional: fetch metadata from the old vCenter first (or use an existing -MetaDir) ---
if ($SourceServer) {
    if (-not $SourcePassword) { throw "-SourceServer requires -SourcePassword" }
    if (-not $MetaDir) { $MetaDir = Join-Path (Get-Location) ('meta-{0}-{1:yyyyMMdd-HHmmss}' -f ($SourceServer -replace '[^\w.-]','_'), (Get-Date)) }
    Write-Host "==================== Exporting folders / tags / attributes / notes from old vCenter $SourceServer ===================="
    $exp = @{ Server = $SourceServer; User = $SourceUser; Password = $SourcePassword; OutDir = $MetaDir }
    if ($SourceDatacenter) { $exp.Datacenter = $SourceDatacenter }
    & "$here\Export-VcMeta.ps1" @exp
    Write-Host ""
}
if ($MetaDir) {
    $MetaDir = (Resolve-Path $MetaDir).Path
    if (-not $PlacementCsv) { $PlacementCsv = Join-Path $MetaDir 'vm-placement.csv' }
    $CreateFolders = $true
}

# --- Connect ---
if (-not $Credential -and $User) {
    $sec = ConvertTo-SecureString $Password -AsPlainText -Force
    $Credential = New-Object System.Management.Automation.PSCredential($User, $sec)
}
Write-Host "[*] Connecting to $Server ..."
$vc = Connect-VIServer -Server $Server -Credential $Credential
Write-Host "[+] Connected to $($vc.Name)  ($($vc.Version) build $($vc.Build))"
if ($DryRun) { Write-Host "[!] DryRun: scan only, nothing is registered" -ForegroundColor Yellow }

# --- vmx paths already in the inventory (skipped) ---
$registered = @{}
foreach ($v in (Get-View -ViewType VirtualMachine -Property Config.Files.VmPathName -Server $vc)) {
    $p = $v.Config.Files.VmPathName
    if ($p) { $registered[$p.ToLower()] = $true }
}
Write-Host "[*] Inventory currently has $($registered.Count) VM(s)/template(s)"

# --- Folder resolution (same rules as Import-VcMeta) ---
$script:FolderCache = @{}
function Resolve-VmFolder {
    param($Dc, [string]$Path, [switch]$Create)
    $root = Get-Folder -Id $Dc.ExtensionData.VmFolder.ToString() -Server $vc
    if ([string]::IsNullOrEmpty($Path)) { return $root }
    $key = "$($Dc.Name)|$Path"
    if ($script:FolderCache.ContainsKey($key)) { return $script:FolderCache[$key] }
    $parent = $root
    $acc = @()
    foreach ($seg in ($Path -split '/')) {
        $acc += $seg
        $subKey = "$($Dc.Name)|$($acc -join '/')"
        if ($script:FolderCache.ContainsKey($subKey)) { $parent = $script:FolderCache[$subKey]; continue }
        $child = Get-Folder -Location $parent -Name $seg -NoRecursion -Server $vc -ErrorAction SilentlyContinue | Select-Object -First 1
        if (-not $child) {
            if (-not $Create) { return $null }
            if ($DryRun) { Add-Result 'WouldCreateFolder' $subKey; return $null }
            $child = New-Folder -Name $seg -Location $parent -Server $vc
            Add-Result 'CreatedFolder' $subKey
        }
        $script:FolderCache[$subKey] = $child
        $parent = $child
    }
    return $parent
}

# --- Read displayName from the vmx (RegisterVM's name must not be empty: PowerCLI sends $null as '' and vCenter rejects it) ---
function Get-VmxDisplayName {
    param([string]$DcName, [string]$DsName, [string]$RelPath)   # RelPath e.g. dir/vm.vmx
    $tmp = Join-Path ([IO.Path]::GetTempPath()) ("vmx-" + [guid]::NewGuid().ToString('N') + ".vmx")
    try {
        Copy-DatastoreItem -Item ("vmstore:\{0}\{1}\{2}" -f $DcName, $DsName, ($RelPath -replace '/', '\')) -Destination $tmp -Force -ErrorAction Stop | Out-Null
        $m = Select-String -Path $tmp -Pattern '^\s*displayName\s*=\s*"(.*)"\s*$' | Select-Object -First 1
        if ($m) { return $m.Matches[0].Groups[1].Value }
    } catch { }
    finally { Remove-Item $tmp -Force -ErrorAction SilentlyContinue }
    return $null
}

# --- Placement map (VMName -> FolderPath) ---
$placement = @{}
if ($PlacementCsv) {
    foreach ($r in (Import-Csv $PlacementCsv)) {
        if ($r.FolderPath) { $placement[$r.VMName] = $r.FolderPath }
    }
    Write-Host "[*] Placement map: $($placement.Count) row(s) ($PlacementCsv)"
}

# =====================================================================
# Mode B: register each row of unregistered.csv (main path)
# =====================================================================
$csvRows = @()
if ($UnregisteredCsv) {
    $UnregisteredCsv = (Resolve-Path $UnregisteredCsv).Path
    $csvRows = @(Import-Csv $UnregisteredCsv)
    if (-not $csvRows.Count) { throw "$UnregisteredCsv is empty" }
    foreach ($col in 'VMName','VmPathName') { if ($csvRows[0].PSObject.Properties.Name -notcontains $col) { throw "$UnregisteredCsv is missing column $col" } }
    $hasFolder = $csvRows[0].PSObject.Properties.Name -contains 'FolderPath'
    $CreateFolders = $true
    Write-Host "`n=== Registering $($csvRows.Count) VM(s) from $(Split-Path $UnregisteredCsv -Leaf) ==="
    $dsCache = @{}; $hostCache = @{}; $poolCache = @{}; $rr = @{}
    # For reconciliation: the datastores covered by this CSV
    $Datastore = @($csvRows | ForEach-Object { ($_.VmPathName -replace '^\[([^\]]+)\].*$', '$1') } | Sort-Object -Unique)

    foreach ($row in $csvRows) {
        $vmx = $row.VmPathName
        if ([string]::IsNullOrWhiteSpace($vmx)) { Add-Result 'SkippedNoPath' $row.VMName; continue }
        if ($registered.ContainsKey($vmx.ToLower())) { Add-Result 'AlreadyRegistered' $vmx; continue }
        if (-not (Test-NameMatch $row.VMName)) { Add-Result 'Filtered' $vmx; continue }
        $isTpl = ($row.IsTemplate -eq 'True')
        if ($isTpl -and $NoTemplates) { Add-Result 'SkippedTemplate' $vmx; continue }
        $dsName = ($vmx -replace '^\[([^\]]+)\].*$', '$1')

        if (-not $dsCache.ContainsKey($dsName)) {
            $dsObj = Get-Datastore -Name $dsName -Server $vc -ErrorAction SilentlyContinue
            $dsCache[$dsName] = $dsObj
            $hs = @()
            if ($dsObj) {
                $hs = @($dsObj | Get-VMHost -Server $vc | Where-Object { $_.ConnectionState -eq 'Connected' })
                if ($Cluster) { $hs = @($hs | Where-Object { $_.Parent.Name -eq $Cluster }) }
                if ($VMHost)  { $hs = @($hs | Where-Object { $_.Name -eq $VMHost }) }
            }
            $hostCache[$dsName] = $hs; $rr[$dsName] = 0
            foreach ($h in $hs) {
                $poolCache[$h.Name] = if ($ResourcePool) {
                    $rp = Get-ResourcePool -Name $ResourcePool -Location $h.Parent -Server $vc -ErrorAction SilentlyContinue | Select-Object -First 1
                    if (-not $rp) { throw "Resource pool '$ResourcePool' not found under $($h.Parent.Name)" }
                    $rp.ExtensionData.MoRef
                } else { (Get-View $h.ExtensionData.Parent -Property ResourcePool -Server $vc).ResourcePool }
            }
            if ($hs.Count) { Write-Host "  datastore $dsName -> hosts: $($hs.Name -join ', ')" }
        }
        $dsObj = $dsCache[$dsName]; $hs = $hostCache[$dsName]
        if (-not $dsObj) { Add-Result 'DatastoreNotFound' $vmx "datastore '$dsName' not on the new vCenter (not mounted yet?)"; continue }
        if (-not $hs.Count) { Add-Result 'NoHost' $vmx "no usable host mounts $dsName (Cluster=$Cluster VMHost=$VMHost)"; continue }
        $h = $hs[$rr[$dsName] % $hs.Count]; $rr[$dsName]++

        $dc = $dsObj.Datacenter
        $fp = if ($hasFolder) { $row.FolderPath } else { $Folder }
        $dest = Resolve-VmFolder $dc $fp -Create:$CreateFolders
        $kind = if ($isTpl) { 'Template' } else { 'VM' }
        if ($DryRun) {
            [void]$script:RegisteredNames.Add($row.VMName)
            Add-Result 'WouldRegister' $vmx "$kind '$($row.VMName)' -> host=$($h.Name) folder=$(if ($fp) { $fp } else { '(DC root)' })"
            continue
        }
        if (-not $dest) { Add-Result 'FolderNotFound' $vmx $fp; continue }
        try {
            $fv = Get-View $dest.ExtensionData.MoRef -Server $vc
            $pool = if ($isTpl) { $null } else { $poolCache[$h.Name] }
            $vmRef = $fv.RegisterVM($vmx, $row.VMName, $isTpl, $pool, $h.ExtensionData.MoRef)   # registered directly into the target folder
            $registered[$vmx.ToLower()] = $true
            [void]$script:RegisteredNames.Add($row.VMName)
            Add-Result 'Registered' $vmx "$kind '$($row.VMName)' on $($h.Name) -> $(if ($fp) { $fp } else { '(DC root)' })"
        } catch {
            Add-Result 'Failed' $vmx ($_.Exception.Message -split "`n")[0]
        }
    }
    if (-not $PlacementCsv) { $PlacementCsv = $UnregisteredCsv }   # for reconciliation
}

# =====================================================================
# Mode A: scan datastores
# =====================================================================
foreach ($dsName in ($(if ($UnregisteredCsv) { @() } else { $Datastore }))) {
    Write-Host "`n=== Datastore: $dsName ==="
    $ds = Get-Datastore -Name $dsName -Server $vc
    $dc = $ds.Datacenter

    # Usable hosts: mount this datastore and are connected, then filtered by -Cluster / -VMHost
    $hosts = @($ds | Get-VMHost -Server $vc | Where-Object { $_.ConnectionState -eq 'Connected' })
    if ($Cluster) { $hosts = @($hosts | Where-Object { $_.Parent.Name -eq $Cluster }) }
    if ($VMHost)  { $hosts = @($hosts | Where-Object { $_.Name -eq $VMHost }) }
    if (-not $hosts) { Write-Host "  [!] No usable host mounts $dsName (Cluster=$Cluster VMHost=$VMHost), skipped"; Add-Result 'NoHost' $dsName; continue }
    Write-Host "  hosts: $($hosts.Name -join ', ')"

    # resource pool
    $poolByHost = @{}
    foreach ($h in $hosts) {
        if ($ResourcePool) {
            $rp = Get-ResourcePool -Name $ResourcePool -Location $h.Parent -Server $vc -ErrorAction SilentlyContinue | Select-Object -First 1
            if (-not $rp) { throw "Resource pool '$ResourcePool' not found under $($h.Parent.Name)" }
            $poolByHost[$h.Name] = $rp.ExtensionData.MoRef
        } else {
            $poolByHost[$h.Name] = (Get-View $h.ExtensionData.Parent -Property ResourcePool -Server $vc).ResourcePool
        }
    }

    # Target folder
    $targetFolder = Resolve-VmFolder $dc $Folder -Create:$CreateFolders
    if (-not $targetFolder) {
        if ($DryRun) { $targetFolder = Get-Folder -Id $dc.ExtensionData.VmFolder.ToString() -Server $vc }
        else { throw "Folder '$Folder' does not exist (use -CreateFolders to create it)" }
    }
    $folderView = Get-View $targetFolder.ExtensionData.MoRef -Server $vc

    # Scan the datastore
    $browser = Get-View $ds.ExtensionData.Browser -Server $vc
    $spec = New-Object VMware.Vim.HostDatastoreBrowserSearchSpec
    $spec.MatchPattern = @('*.vmx', '*.vmtx')
    $spec.Details = New-Object VMware.Vim.FileQueryFlags
    $spec.Details.FileType = $true
    $spec.Details.FileSize = $true
    $spec.Details.Modification = $true
    $taskRef = $browser.SearchDatastoreSubFolders_Task("[$($ds.Name)]", $spec)
    $task = Get-View $taskRef -Server $vc
    while ($task.Info.State -in 'queued','running') { Start-Sleep -Milliseconds 500; $task.UpdateViewData('Info') }
    if ($task.Info.State -ne 'success') { throw "Scan of $dsName failed: $($task.Info.Error.LocalizedMessage)" }

    $found = New-Object System.Collections.ArrayList
    foreach ($r in $task.Info.Result) {
        foreach ($f in $r.File) {
            $ext = [IO.Path]::GetExtension($f.Path).ToLower()
            if ($ext -notin '.vmx', '.vmtx') { continue }
            # FolderPath is "[ds]" for the root or "[ds] dir/" for a sub-directory; build "[ds] dir/file.vmx" (same format as VmPathName)
            $dir = $r.FolderPath.Trim()
            $full = if ($dir.EndsWith(']')) { "$dir $($f.Path)" } else { $dir.TrimEnd('/') + '/' + $f.Path }
            [void]$found.Add([pscustomobject]@{
                Path       = $full
                Name       = [IO.Path]::GetFileNameWithoutExtension($f.Path)
                IsTemplate = ($ext -eq '.vmtx')
                Size       = $f.FileSize
            })
        }
    }
    Write-Host "  found $($found.Count) vmx/vmtx file(s)"

    $i = 0
    foreach ($item in ($found | Sort-Object Path)) {
        if ($registered.ContainsKey($item.Path.ToLower())) { Add-Result 'AlreadyRegistered' $item.Path; continue }
        if (-not (Test-NameMatch $item.Name))              { Add-Result 'Filtered' $item.Path; continue }
        if ($item.IsTemplate -and $NoTemplates)            { Add-Result 'SkippedTemplate' $item.Path; continue }

        $h = $hosts[$i % $hosts.Count]; $i++
        $kind = if ($item.IsTemplate) { 'Template' } else { 'VM' }

        # VM name: displayName from the vmx by default (RegisterVM's name must not be empty), file name as fallback
        $name = $null
        if (-not $NameFromFile) {
            $rel = ($item.Path -replace '^\[[^\]]+\]\s*', '')
            $name = Get-VmxDisplayName $dc.Name $ds.Name $rel
        }
        if ([string]::IsNullOrWhiteSpace($name)) { $name = $item.Name }

        if ($DryRun) {
            [void]$script:RegisteredNames.Add($name)
            $where = if ($placement.ContainsKey($name)) { $placement[$name] } else { $targetFolder.Name }
            Add-Result 'WouldRegister' $item.Path "$kind '$name' -> host=$($h.Name) folder=$where"
            continue
        }

        try {
            $pool = if ($item.IsTemplate) { $null } else { $poolByHost[$h.Name] }
            $vmRef = $folderView.RegisterVM($item.Path, $name, $item.IsTemplate, $pool, $h.ExtensionData.MoRef)
            $registered[$item.Path.ToLower()] = $true
            $vmObj = Get-VIObjectByVIView -MORef $vmRef -Server $vc
            [void]$script:RegisteredNames.Add($vmObj.Name)
            Add-Result 'Registered' $item.Path "$kind '$($vmObj.Name)' on $($h.Name)"

            # Move into its folder according to the placement map
            if ($placement.ContainsKey($vmObj.Name)) {
                $dest = Resolve-VmFolder $dc $placement[$vmObj.Name] -Create:$CreateFolders
                if (-not $dest) { Add-Result 'FolderNotFound' $vmObj.Name $placement[$vmObj.Name] }
                elseif ($dest.Id -ne $targetFolder.Id) {
                    if ($item.IsTemplate) { $null = Move-Template -Template $vmObj -Destination $dest -Server $vc -Confirm:$false }
                    else                  { $null = Move-VM -VM $vmObj -InventoryLocation $dest -Server $vc -Confirm:$false }
                    Add-Result 'Placed' $vmObj.Name $placement[$vmObj.Name]
                }
            }
        } catch {
            Add-Result 'Failed' $item.Path ($_.Exception.Message -split "`n")[0]
        }
    }
}

# --- Optional: apply attribute values / notes / tags from the old vCenter (folders were handled at registration) ---
if ($MetaDir -and $script:RegisteredNames.Count) {
    Write-Host "`n==================== Applying attribute values / notes / tags from the old vCenter ($($script:RegisteredNames.Count) VM(s)) ===================="
    $imp = @{
        Server = $Server; InDir = $MetaDir
        Include = @('CustomAttributes','Notes','Tags')
        OnlyVMs = @($script:RegisteredNames)
        ReportPath = ($ReportPath -replace '\.csv$', '-meta.csv')
    }
    if ($Credential) { $imp.Credential = $Credential } else { $imp.User = $User; $imp.Password = $Password }
    if ($DatacenterMap) { $imp.DatacenterMap = $DatacenterMap }
    if ($DryRun)        { $imp.DryRun = $true }
    & "$here\Import-VcMeta.ps1" @imp
}

# --- Reconciliation: VMs in the CSV / manifest on these datastores that are still not on the new vCenter ---
if ($PlacementCsv -and (Test-Path $PlacementCsv)) {
    $manifest = @(Import-Csv $PlacementCsv)
    if ($manifest.Count -and ($manifest[0].PSObject.Properties.Name -contains 'VmPathName')) {
        $onTarget = @{}
        foreach ($v in (Get-View -ViewType VirtualMachine -Property Name -Server $vc)) { $onTarget[$v.Name] = $true }
        if (-not $DryRun) { foreach ($n in $script:RegisteredNames) { $onTarget[$n] = $true } }
        # Expected set: prefer unregistered.csv written by Unregister-VmFromOldVc (what was actually removed);
        # otherwise the manifest rows whose vmx is on these datastores
        $unregLog = Join-Path (Split-Path $PlacementCsv -Parent) 'unregistered.csv'
        if (Test-Path $unregLog) {
            $expected = @(Import-Csv $unregLog | Where-Object { $p = $_.VmPathName; $p -and ($Datastore | Where-Object { $p.StartsWith("[$_] ") }) })
            $basis = 'unregistered.csv'
        } else {
            $expected = @($manifest | Where-Object { $p = $_.VmPathName; $p -and ($Datastore | Where-Object { $p.StartsWith("[$_] ") }) })
            $basis = 'vm-placement.csv'
        }
        $pending = @($expected | Where-Object { -not $onTarget.ContainsKey($_.VMName) })
        Write-Host "`n================ Reconciliation ($basis vs new vCenter) ================"
        if ($DryRun) {
            $covered = @($pending | Where-Object { $script:RegisteredNames -contains $_.VMName }).Count
            Write-Host ("  VM(s) from the old vCenter on {0}: {1}; already on the new vCenter: {2}; still pending: {3} (this run would register {4}, leaving {5})" -f ($Datastore -join ","), $expected.Count, ($expected.Count - $pending.Count), $pending.Count, $covered, ($pending.Count - $covered))
        } else {
            Write-Host ("  VM(s) from the old vCenter on {0}: {1}; already on the new vCenter: {2}; still pending: {3}" -f ($Datastore -join ","), $expected.Count, ($expected.Count - $pending.Count), $pending.Count)
        }
        foreach ($m in $pending) {
            Add-Result 'PendingOnSource' $m.VMName "$($m.VmPathName)"
        }
    }
}

# --- Summary ---
Write-Host "`n================ Result ================"
foreach ($k in ($script:Stats.Keys | Sort-Object)) { Write-Host ("  {0,-20} {1,5}" -f $k, $script:Stats[$k]) }
if ($script:Report.Count) {
    $script:Report | Export-Csv -Path $ReportPath -NoTypeInformation -Encoding $enc
    Write-Host "`n[+] Detail report -> $ReportPath"
}
Disconnect-VIServer -Server $vc -Confirm:$false | Out-Null
