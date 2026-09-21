#Requires -Version 7.0
<#
.SYNOPSIS
  Old vCenter side (step 3): remove VMs / templates from the inventory according
  to the export manifest (vm-placement.csv). The files stay on the datastore.

.DESCRIPTION
  Produces <MetaDir>\unregistered.csv (name / vmx path / original folder / time).
  That file is the input for the next step, Register-VmxFromDatastore -UnregisteredCsv,
  and can be edited in Excel first (e.g. to change the target folder).

  Safety rules:
    - Refuses to run without -MetaDir (Export-VcMeta must have run first):
      once a VM is unregistered its tag assignments and attribute values are gone.
    - The vmx path in the manifest must match the current one (stale export = skipped).
    - Powered-on VMs are skipped unless -ShutdownFirst is given.
    - With -Datastore, VMs that span datastores are reported: a VM in the list with a
      vmdk on another datastore (WarnDiskElsewhere) and a VM outside the list with a
      vmdk on this datastore (WarnDiskOnDatastore). Warning only, not blocked.
    - "All VMs" is not accepted: at least one filter is required.

  Filters can be combined; when several are given the result is the INTERSECTION
  (e.g. -Cluster cl01 -Datastore ds01 = VMs in cl01 whose vmx is on ds01).

  Self-signed vCenter certificates are accepted automatically. The PowerCLI CEIP
  prompt is suppressed so the script never blocks on first run.

.EXAMPLE
  # All VMs in the manifest whose vmx is on ds01 (dry run first)
  .\Unregister-VmFromOldVc.ps1 -Server vc-old.example.com -Password '<pw>' -MetaDir .\export-A -Datastore ds01 -DryRun

.EXAMPLE
  # VMs in cluster cl01 AND on ds01
  .\Unregister-VmFromOldVc.ps1 -Server vc-old.example.com -Password '<pw>' -MetaDir .\export-A -Cluster cl01 -Datastore ds01

.EXAMPLE
  # Only one folder subtree
  .\Unregister-VmFromOldVc.ps1 -Server vc-old.example.com -Password '<pw>' -MetaDir .\export-A -Folder 'Linux'
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$Server,
    [string]$User = 'administrator@vsphere.local',
    [Parameter(Mandatory)][string]$Password,
    [Parameter(Mandatory)][string]$MetaDir,        # -OutDir of Export-VcMeta

    # Filters: any combination, intersection when several are given
    [string[]]$Datastore,              # manifest rows whose vmx is on these datastores
    [string[]]$Cluster,                # VMs currently running on these clusters (queried live on the old vCenter)
    [string[]]$VMHost,                 # VMs currently running on these hosts (queried live)
    [string[]]$Folder,                 # manifest rows inside these folder subtrees
    [string[]]$VM,                     # explicit VM names
    [switch]$ShutdownFirst,
    [int]$ShutdownTimeoutSec = 300,
    [switch]$DryRun,
    [string]$ReportPath
)

$ErrorActionPreference = 'Stop'
$enc = 'utf8BOM'   # PowerShell 7: explicit UTF-8 with BOM so Excel and Windows PowerShell read the CSV correctly
if (-not $Datastore -and -not $Cluster -and -not $VMHost -and -not $Folder -and -not $VM) { throw "Give at least one of -Datastore / -Cluster / -VMHost / -Folder / -VM. 'All VMs' is not accepted." }
$MetaDir = (Resolve-Path $MetaDir).Path
$manifestPath = Join-Path $MetaDir 'vm-placement.csv'
if (-not (Test-Path $manifestPath)) { throw "$MetaDir has no vm-placement.csv -- run Export-VcMeta.ps1 first" }
$manifest = @(Import-Csv $manifestPath)
if (-not $manifest.Count -or ($manifest[0].PSObject.Properties.Name -notcontains 'VmPathName')) { throw "vm-placement.csv is too old (no VmPathName column); re-export with the current Export-VcMeta" }
if (-not $ReportPath) { $ReportPath = Join-Path $MetaDir ('unregister-report-{0:yyyyMMdd-HHmmss}.csv' -f (Get-Date)) }
$logPath = Join-Path $MetaDir 'unregistered.csv'

$script:Report = New-Object System.Collections.ArrayList
$script:Stats  = @{}
function Add-Result {
    param([string]$Action,[string]$Target,[string]$Detail = '')
    [void]$script:Report.Add([pscustomobject]@{ Time = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'); Action = $Action; Target = $Target; Detail = $Detail })
    if (-not $script:Stats.ContainsKey($Action)) { $script:Stats[$Action] = 0 }
    $script:Stats[$Action]++
    $tag = switch -Wildcard ($Action) { 'Unregistered' { '[+]' } 'Would*' { '[~]' } 'Failed*' { '[X]' } 'Skipped*' { '[-]' } 'Warn*' { '[!]' } default { '[ ]' } }
    Write-Host ("  {0} {1,-20} {2}  {3}" -f $tag, $Action, $Target, $Detail)
}

# --- PowerCLI: accept self-signed certificates, suppress CEIP prompt ---
try { Import-Module VMware.VimAutomation.Core -ErrorAction Stop } catch { throw "VMware.PowerCLI is not installed. Run: Install-Module VMware.PowerCLI -Scope CurrentUser" }
try { Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -ParticipateInCeip $false -DisplayDeprecationWarnings $false -Scope User -Confirm:$false | Out-Null } catch { }
Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -DisplayDeprecationWarnings $false -Scope Session -Confirm:$false | Out-Null

# --- Connect (-Cluster / -VMHost need a live query on the old vCenter) ---
$vc = Connect-VIServer -Server $Server -User $User -Password $Password
Write-Host "[+] Connected to $($vc.Name) ($($vc.Version))"
if ($DryRun) { Write-Host "[!] DryRun: check only, nothing is unregistered" -ForegroundColor Yellow }

# --- Select manifest rows (intersection of all given filters) ---
$inCluster = $null
if ($Cluster -or $VMHost) {
    $inCluster = @{}
    $hostIds = @{}
    foreach ($n in $Cluster) {
        $o = Get-Cluster -Name $n -Server $vc -ErrorAction SilentlyContinue
        if (-not $o) { Write-Host "  [!] Cluster '$n' not found on the old vCenter"; continue }
        foreach ($h in ($o | Get-VMHost -Server $vc)) { $hostIds[$h.ExtensionData.MoRef.ToString()] = $true }
    }
    foreach ($n in $VMHost) {
        $o = Get-VMHost -Name $n -Server $vc -ErrorAction SilentlyContinue
        if (-not $o) { Write-Host "  [!] Host '$n' not found on the old vCenter"; continue }
        $hostIds[$o.ExtensionData.MoRef.ToString()] = $true
    }
    # Filter via views on Runtime.Host: covers VMs and templates in one pass (Get-Template -Location does not accept a cluster)
    foreach ($v in (Get-View -ViewType VirtualMachine -Property Name,Runtime.Host -Server $vc)) {
        if ($v.Runtime.Host -and $hostIds.ContainsKey($v.Runtime.Host.ToString())) { $inCluster[$v.Name] = $true }
    }
    Write-Host "[*] Cluster/host $((@($Cluster) + @($VMHost) | Where-Object { $_ }) -join ',') currently has $($inCluster.Count) VM(s)"
}
$picked = @($manifest | Where-Object {
    $r = $_
    if ($VM        -and ($VM -notcontains $r.VMName)) { return $false }
    if ($Datastore -and -not ($Datastore | Where-Object { $r.VmPathName.StartsWith("[$_] ") })) { return $false }
    if ($Folder    -and -not ($Folder | Where-Object { $f = $_.Trim('/'); $r.FolderPath -eq $f -or $r.FolderPath -like "$f/*" })) { return $false }
    if ($null -ne $inCluster -and -not $inCluster.ContainsKey($r.VMName)) { return $false }
    $true
})
$cond = @()
if ($VM) { $cond += "VM=$($VM -join ',')" }; if ($Datastore) { $cond += "datastore=$($Datastore -join ',')" }
if ($Cluster) { $cond += "cluster=$($Cluster -join ',')" }; if ($VMHost) { $cond += "host=$($VMHost -join ',')" }; if ($Folder) { $cond += "folder=$($Folder -join ',')" }
Write-Host "[*] Manifest has $($manifest.Count) VM(s); $($picked.Count) match [$($cond -join ' AND ')]"
if (-not $picked.Count) { Disconnect-VIServer -Server $vc -Confirm:$false | Out-Null; return }

# --- Cross-datastore check (meaningful only with -Datastore) ---
# When a single datastore/LUN is moved on its own, these two kinds of VM break:
#   A) in the list (vmx on the target datastore) but with a vmdk on another datastore -> disk missing after registration
#   B) not in the list (vmx elsewhere) but with a vmdk on the target datastore        -> a VM left behind loses a disk
if ($Datastore) {
    $pickedNames = @{}; foreach ($r in $picked) { $pickedNames[$r.VMName] = $true }
    foreach ($v in (Get-View -ViewType VirtualMachine -Property Name,Config.Files.VmPathName,Config.Hardware.Device -Server $vc)) {
        $homeDs = ($v.Config.Files.VmPathName -replace '^\[([^\]]+)\].*$', '$1')
        $diskDs = @($v.Config.Hardware.Device | Where-Object { $_ -is [VMware.Vim.VirtualDisk] -and $_.Backing.FileName } |
                    ForEach-Object { ($_.Backing.FileName -replace '^\[([^\]]+)\].*$', '$1') } | Sort-Object -Unique)
        $homeIn  = $Datastore -contains $homeDs
        $diskOut = @($diskDs | Where-Object { $Datastore -notcontains $_ })
        $diskIn  = @($diskDs | Where-Object { $Datastore -contains $_ })
        if ($homeIn -and $diskOut.Count -and $pickedNames.ContainsKey($v.Name)) {
            Add-Result 'WarnDiskElsewhere' $v.Name "has vmdk on $($diskOut -join ','); moving only $($Datastore -join ',') leaves it without that disk"
        }
        if (-not $homeIn -and $diskIn.Count) {
            Add-Result 'WarnDiskOnDatastore' $v.Name "vmx is on $homeDs but a vmdk is on $($diskIn -join ','); not in the list yet its disk would move"
        }
    }
}

$done = New-Object System.Collections.ArrayList
foreach ($r in $picked) {
    $o = Get-VM -Name $r.VMName -Server $vc -ErrorAction SilentlyContinue
    $isTpl = $false
    if (-not $o) { $o = Get-Template -Name $r.VMName -Server $vc -ErrorAction SilentlyContinue; $isTpl = [bool]$o }
    if (-not $o) { Add-Result 'SkippedNotFound' $r.VMName 'not on the old vCenter any more (already unregistered?)'; continue }
    if (@($o).Count -gt 1) { Add-Result 'SkippedAmbiguous' $r.VMName 'several objects with this name; narrow down by folder or UUID'; continue }

    $curVmx = $o.ExtensionData.Config.Files.VmPathName
    if ($curVmx -ne $r.VmPathName) { Add-Result 'SkippedStale' $r.VMName "manifest $($r.VmPathName) != current $curVmx; re-export first"; continue }

    if (-not $isTpl -and $o.PowerState -ne 'PoweredOff') {
        if (-not $ShutdownFirst) { Add-Result 'SkippedPoweredOn' $r.VMName "$($o.PowerState); power off first or use -ShutdownFirst"; continue }
        if ($DryRun) { Add-Result 'WouldShutdown' $r.VMName }
        else {
            try { $null = Shutdown-VMGuest -VM $o -Confirm:$false -Server $vc } catch { Add-Result 'FailedShutdown' $r.VMName ($_.Exception.Message -split "`n")[0]; continue }
            $deadline = (Get-Date).AddSeconds($ShutdownTimeoutSec)
            do { Start-Sleep 5; $o = Get-VM -Name $r.VMName -Server $vc } while ($o.PowerState -ne 'PoweredOff' -and (Get-Date) -lt $deadline)
            if ($o.PowerState -ne 'PoweredOff') { Add-Result 'FailedShutdown' $r.VMName "still not off after $ShutdownTimeoutSec s"; continue }
            Add-Result 'Shutdown' $r.VMName
        }
    }

    if ($DryRun) { Add-Result 'WouldUnregister' $r.VMName $curVmx; continue }
    try {
        if ($isTpl) { Remove-Template -Template $o -Confirm:$false -Server $vc } else { Remove-VM -VM $o -Confirm:$false -Server $vc }
        Add-Result 'Unregistered' $r.VMName $curVmx
        [void]$done.Add([pscustomobject]@{
            Time = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'); VMName = $r.VMName; IsTemplate = $isTpl
            VmPathName = $curVmx; Datacenter = $r.Datacenter; FolderPath = $r.FolderPath
            InstanceUuid = $r.InstanceUuid; VMHost = $r.VMHost; Cluster = $r.Cluster; Server = $Server
        })
    } catch {
        Add-Result 'FailedUnregister' $r.VMName ($_.Exception.Message -split "`n")[0]
    }
}

if ($done.Count) {
    if (Test-Path $logPath) { $done | Export-Csv -Path $logPath -NoTypeInformation -Encoding $enc -Append }
    else { $done | Export-Csv -Path $logPath -NoTypeInformation -Encoding $enc }
    Write-Host "`n[+] $($done.Count) unregistered VM(s) recorded in $logPath (input for Register-VmxFromDatastore -UnregisteredCsv)"
}
Write-Host "`n================ Result ================"
foreach ($k in ($script:Stats.Keys | Sort-Object)) { Write-Host ("  {0,-22} {1,5}" -f $k, $script:Stats[$k]) }
$script:Report | Export-Csv -Path $ReportPath -NoTypeInformation -Encoding $enc
Write-Host "[+] Detail report -> $ReportPath"
Disconnect-VIServer -Server $vc -Confirm:$false | Out-Null
