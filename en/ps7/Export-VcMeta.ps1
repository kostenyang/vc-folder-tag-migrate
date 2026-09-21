#Requires -Version 7.0
<#
.SYNOPSIS
  Export folder tree / VM placement / tags / custom attributes / notes from the
  source vCenter into CSV files.

.DESCRIPTION
  Step 0 of the migration runbook. Produces 8 CSV files in -OutDir:
    folders.csv, vm-placement.csv, notes.csv, tag-categories.csv, tags.csv,
    tag-assignments.csv, custom-attributes.csv, custom-attribute-values.csv
  The CSV files are UTF-8 (with BOM) and can be edited in Excel before import.

  Self-signed vCenter certificates are accepted automatically. The PowerCLI CEIP
  prompt is suppressed so the script never blocks on first run.

.EXAMPLE
  .\Export-VcMeta.ps1 -Server vc-old.example.com -User administrator@vsphere.local -Password '<pw>' -OutDir .\export-A

.EXAMPLE
  # Only one VM folder subtree (and the VMs inside it)
  .\Export-VcMeta.ps1 -Server vc-old.example.com -User administrator@vsphere.local -Password '<pw>' -OutDir .\export-A -Folder 'Linux'
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$Server,
    [string]$User,
    [string]$Password,
    [pscredential]$Credential,
    [string]$OutDir = "./vc-meta-export",
    [string[]]$Datacenter,             # only these datacenters
    [string[]]$Folder,                 # only these VM folder subtrees and the VMs inside, e.g. -Folder 'Linux','MGMT/Prod'
    [switch]$AllDefinitions,           # with -Folder: also export tag/attribute definitions that are not used in scope
    [ValidateSet('Folders','VMPlacement','Tags','CustomAttributes','Notes')]
    [string[]]$Include = @('Folders','VMPlacement','Tags','CustomAttributes','Notes')
)

$ErrorActionPreference = 'Stop'
$enc = 'utf8BOM'   # PowerShell 7: explicit UTF-8 with BOM so Excel and Windows PowerShell read the CSV correctly

# The vCenter tagging service (vAPI/CIS) occasionally returns 503 right after a new session; retry a few times.
function Invoke-WithRetry {
    param([scriptblock]$Script, [string]$What = 'call', [int]$Times = 3, [int]$DelaySec = 10)
    for ($i = 1; $i -le $Times; $i++) {
        try { return (& $Script) }
        catch {
            if ($i -eq $Times) { throw }
            Write-Host "  [!] $What failed ($(($_.Exception.Message -split "`n")[0])), retrying in $DelaySec s ($i/$($Times - 1))"
            Start-Sleep -Seconds $DelaySec
        }
    }
}
function Write-Meta {
    param($Rows, [string]$Path, [string[]]$Columns)
    if (-not $Rows -or @($Rows).Count -eq 0) {
        Set-Content -Path $Path -Value (($Columns | ForEach-Object { '"' + $_ + '"' }) -join ',') -Encoding $enc
        Write-Host ("  -> {0,-30} {1,5} rows" -f (Split-Path $Path -Leaf), 0)
    } else {
        @($Rows) | Select-Object $Columns | Export-Csv -Path $Path -NoTypeInformation -Encoding $enc
        Write-Host ("  -> {0,-30} {1,5} rows" -f (Split-Path $Path -Leaf), @($Rows).Count)
    }
}

# --- PowerCLI: accept self-signed certificates, suppress CEIP prompt ---
try { Import-Module VMware.VimAutomation.Core -ErrorAction Stop } catch { throw "VMware.PowerCLI is not installed. Run: Install-Module VMware.PowerCLI -Scope CurrentUser" }
try { Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -ParticipateInCeip $false -DisplayDeprecationWarnings $false -Scope User -Confirm:$false | Out-Null } catch { }
Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -DisplayDeprecationWarnings $false -Scope Session -Confirm:$false | Out-Null

# --- Connect ---
if (-not $Credential -and $User) {
    $sec = ConvertTo-SecureString $Password -AsPlainText -Force
    $Credential = New-Object System.Management.Automation.PSCredential($User, $sec)
}
Write-Host "[*] Connecting to $Server ..."
$vc = Connect-VIServer -Server $Server -Credential $Credential
Write-Host "[+] Connected to $($vc.Name)  ($($vc.Version) build $($vc.Build))"

New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$OutDir = (Resolve-Path $OutDir).Path

# --- Datacenter / folder index ---
$dcViews = Get-View -ViewType Datacenter -Property Name,VmFolder,HostFolder,DatastoreFolder,NetworkFolder -Server $vc
if ($Datacenter) { $dcViews = $dcViews | Where-Object { $Datacenter -contains $_.Name } }
if (-not $dcViews) { throw "No datacenter found (filter: $($Datacenter -join ','))" }

$folderViews = Get-View -ViewType Folder -Property Name,Parent,ChildType -Server $vc
$fById = @{}
foreach ($f in $folderViews) { $fById[$f.MoRef.ToString()] = $f }

# root folder moref -> @{ dc = <name>; type = VM|HostAndCluster|Datastore|Network }
$rootInfo = @{}
foreach ($dc in $dcViews) {
    $rootInfo[$dc.VmFolder.ToString()]        = @{ dc = $dc.Name; type = 'VM' }
    $rootInfo[$dc.HostFolder.ToString()]      = @{ dc = $dc.Name; type = 'HostAndCluster' }
    $rootInfo[$dc.DatastoreFolder.ToString()] = @{ dc = $dc.Name; type = 'Datastore' }
    $rootInfo[$dc.NetworkFolder.ToString()]   = @{ dc = $dc.Name; type = 'Network' }
}

# folder moref -> @{ dc; type; path; name }   (path is relative to the root folder; root itself has path = '')
$folderInfo = @{}
foreach ($f in $folderViews) {
    $id = $f.MoRef.ToString()
    if ($rootInfo.ContainsKey($id)) {
        $folderInfo[$id] = @{ dc = $rootInfo[$id].dc; type = $rootInfo[$id].type; path = ''; name = $f.Name }
        continue
    }
    $names = New-Object System.Collections.ArrayList
    $cur = $f
    $found = $null
    while ($cur) {
        [void]$names.Add($cur.Name)
        $parentId = if ($cur.Parent) { $cur.Parent.ToString() } else { $null }
        if (-not $parentId) { break }
        if ($rootInfo.ContainsKey($parentId)) { $found = $rootInfo[$parentId]; break }
        if (-not $fById.ContainsKey($parentId)) { break }   # parent is not a folder (vApp / datacenter)
        $cur = $fById[$parentId]
    }
    if ($found) {
        $names.Reverse()
        $folderInfo[$id] = @{ dc = $found.dc; type = $found.type; path = ($names -join '/'); name = $f.Name }
    }
}

# --- Folder scope (-Folder) ---
# $scopeFolderIds / $scopeVmIds = $null means "no scope, whole vCenter"
$scopeFolderIds = $null
$scopeVmIds     = $null
if ($Folder) {
    $scopeFolderIds = @{}
    foreach ($kv in $folderInfo.GetEnumerator()) {
        if ($kv.Value.type -ne 'VM' -or [string]::IsNullOrEmpty($kv.Value.path)) { continue }
        foreach ($f in $Folder) {
            $fn = $f.Trim('/')
            if ($kv.Value.path -eq $fn -or $kv.Value.path -like "$fn/*") { $scopeFolderIds[$kv.Key] = $true; break }
        }
    }
    if ($scopeFolderIds.Count -eq 0) { throw "No matching VM folder: $($Folder -join ', ')" }
    Write-Host "[*] Scope: $($Folder -join ', ')  -> $($scopeFolderIds.Count) folder(s) incl. subtree"
}

Write-Host "[*] Exporting ..."

# --- 1. Folders ---
if ($Include -contains 'Folders') {
    $rows = foreach ($kv in $folderInfo.GetEnumerator()) {
        if ([string]::IsNullOrEmpty($kv.Value.path)) { continue }   # root folders are not exported
        if ($scopeFolderIds -and -not $scopeFolderIds.ContainsKey($kv.Key)) { continue }
        [pscustomobject]@{
            Datacenter = $kv.Value.dc
            FolderType = $kv.Value.type
            Path       = $kv.Value.path
            Name       = $kv.Value.name
            Depth      = ($kv.Value.path -split '/').Count
            MoRef      = $kv.Key
        }
    }
    $rows = $rows | Sort-Object Datacenter, FolderType, Depth, Path
    Write-Meta $rows (Join-Path $OutDir 'folders.csv') @('Datacenter','FolderType','Path','Name','Depth','MoRef')
}

# --- 2. VM placement + 5. Notes ---
if (($Include -contains 'VMPlacement') -or ($Include -contains 'Notes') -or $Folder) {
    if ($Folder) { $scopeVmIds = @{} }
    $vmViews = Get-View -ViewType VirtualMachine -Property Name,Parent,Config.InstanceUuid,Config.Uuid,Config.Annotation,Config.Template,Config.Files.VmPathName,Runtime.PowerState,Runtime.Host -Server $vc
    # host -> name / cluster (recorded in the manifest so Unregister can filter by cluster)
    $hostName = @{}; $hostCluster = @{}
    $clusterName = @{}
    foreach ($cl in (Get-View -ViewType ClusterComputeResource -Property Name -Server $vc)) { $clusterName[$cl.MoRef.ToString()] = $cl.Name }
    foreach ($hv in (Get-View -ViewType HostSystem -Property Name,Parent -Server $vc)) {
        $hostName[$hv.MoRef.ToString()] = $hv.Name
        $pid2 = if ($hv.Parent) { $hv.Parent.ToString() } else { '' }
        $hostCluster[$hv.MoRef.ToString()] = if ($clusterName.ContainsKey($pid2)) { $clusterName[$pid2] } else { '' }
    }
    $place = New-Object System.Collections.ArrayList
    $notes = New-Object System.Collections.ArrayList
    foreach ($vm in $vmViews) {
        $parentId = if ($vm.Parent) { $vm.Parent.ToString() } else { $null }
        $info = if ($parentId -and $folderInfo.ContainsKey($parentId)) { $folderInfo[$parentId] } else { $null }
        $dcName = if ($info) { $info.dc } else { '' }
        if ($Datacenter -and $dcName -and ($Datacenter -notcontains $dcName)) { continue }
        if ($scopeFolderIds) {
            if (-not ($parentId -and $scopeFolderIds.ContainsKey($parentId))) { continue }
            $scopeVmIds[$vm.MoRef.ToString()] = $true
        }
        [void]$place.Add([pscustomobject]@{
            Datacenter   = $dcName
            VMName       = $vm.Name
            InstanceUuid = $vm.Config.InstanceUuid
            BiosUuid     = $vm.Config.Uuid
            IsTemplate   = [bool]$vm.Config.Template
            FolderPath   = if ($info) { $info.path } else { '' }
            InVApp       = (-not $info)
            VmPathName   = $vm.Config.Files.VmPathName
            PowerState   = "$($vm.Runtime.PowerState)"
            VMHost       = if ($vm.Runtime.Host) { $hostName[$vm.Runtime.Host.ToString()] } else { '' }
            Cluster      = if ($vm.Runtime.Host) { $hostCluster[$vm.Runtime.Host.ToString()] } else { '' }
            MoRef        = $vm.MoRef.ToString()
        })
        $ann = $vm.Config.Annotation
        if (-not [string]::IsNullOrWhiteSpace($ann)) {
            [void]$notes.Add([pscustomobject]@{
                Datacenter   = $dcName
                EntityType   = 'VirtualMachine'
                EntityName   = $vm.Name
                InstanceUuid = $vm.Config.InstanceUuid
                Notes        = $ann
            })
        }
    }
    if ($Include -contains 'VMPlacement') {
        Write-Meta ($place | Sort-Object Datacenter,FolderPath,VMName) (Join-Path $OutDir 'vm-placement.csv') @('Datacenter','VMName','InstanceUuid','BiosUuid','IsTemplate','FolderPath','InVApp','VmPathName','PowerState','VMHost','Cluster','MoRef')
    }
    if ($Include -contains 'Notes') {
        Write-Meta ($notes | Sort-Object EntityName) (Join-Path $OutDir 'notes.csv') @('Datacenter','EntityType','EntityName','InstanceUuid','Notes')
    }
}

# --- 3. Tags ---
if ($Include -contains 'Tags') {
    $cats = Invoke-WithRetry { Get-TagCategory -Server $vc } -What 'Get-TagCategory'
    $tags = Invoke-WithRetry { Get-Tag -Server $vc } -What 'Get-Tag'
    # Definitions are written after assignments are known, so that with -Folder only the used categories/tags are kept.

    # Collect entities per type, then query assignments in chunks: a plain Get-TagAssignment
    # aborts entirely when it hits an inaccessible datastore or similar object.
    $targets = New-Object System.Collections.ArrayList
    foreach ($blk in @(
        { Get-VM -Server $vc }, { Get-Template -Server $vc }, { Get-VMHost -Server $vc },
        { Get-Cluster -Server $vc }, { Get-Datacenter -Server $vc },
        { Get-Datastore -Server $vc | Where-Object { $_.State -ne 'Unavailable' } },
        { Get-DatastoreCluster -Server $vc }, { Get-Folder -Server $vc },
        { Get-ResourcePool -Server $vc }, { Get-VDPortgroup -Server $vc }, { Get-VApp -Server $vc }
    )) {
        try { foreach ($o in (& $blk)) { [void]$targets.Add($o) } } catch { Write-Host "  [!] Skipped while collecting entities: $($_.Exception.Message)" }
    }
    $assigns = New-Object System.Collections.ArrayList
    $chunk = 200
    for ($i = 0; $i -lt $targets.Count; $i += $chunk) {
        $slice = $targets[$i..([Math]::Min($i + $chunk - 1, $targets.Count - 1))]
        try {
            foreach ($x in (Get-TagAssignment -Entity $slice -Server $vc)) { [void]$assigns.Add($x) }
        } catch {
            foreach ($one in $slice) {
                try { foreach ($x in (Get-TagAssignment -Entity $one -Server $vc)) { [void]$assigns.Add($x) } }
                catch { Write-Host "  [!] Skipped $($one.Name): $($_.Exception.Message)" }
            }
        }
    }
    $arows = New-Object System.Collections.ArrayList
    foreach ($a in $assigns) {
        $e = $a.Entity
        $moref = $null; $etype = $null
        try { $moref = $e.ExtensionData.MoRef; $etype = $moref.Type } catch { }
        if (-not $etype) { $etype = ($e.GetType().Name -replace 'Impl$','') }
        $id = if ($moref) { $moref.ToString() } else { '' }
        if ($scopeFolderIds -and -not ($scopeFolderIds.ContainsKey($id) -or $scopeVmIds.ContainsKey($id))) { continue }
        $fi = if ($id -and $folderInfo.ContainsKey($id)) { $folderInfo[$id] } else { $null }
        $uuid = ''
        if ($etype -eq 'VirtualMachine') { try { $uuid = $e.ExtensionData.Config.InstanceUuid } catch { } }
        [void]$arows.Add([pscustomobject]@{
            EntityType = $etype
            EntityName = $e.Name
            EntityUuid = $uuid
            Datacenter = if ($fi) { $fi.dc }   else { '' }
            FolderType = if ($fi) { $fi.type } else { '' }
            EntityPath = if ($fi) { $fi.path } else { '' }
            Category   = $a.Tag.Category.Name
            Tag        = $a.Tag.Name
            MoRef      = $id
        })
    }
    # With -Folder, keep only categories/tags that are actually assigned in scope (-AllDefinitions keeps all)
    $usedTag = @{}; $usedCat = @{}
    foreach ($r in $arows) { $usedTag["$($r.Category)|$($r.Tag)"] = $true; $usedCat[$r.Category] = $true }
    $catRows = $cats | Where-Object { -not $scopeFolderIds -or $AllDefinitions -or $usedCat.ContainsKey($_.Name) }
    $tagRows = $tags | Where-Object { -not $scopeFolderIds -or $AllDefinitions -or $usedTag.ContainsKey("$($_.Category.Name)|$($_.Name)") }

    Write-Meta ($catRows | ForEach-Object {
        [pscustomobject]@{
            Name        = $_.Name
            Description = $_.Description
            Cardinality = $_.Cardinality
            EntityType  = ($_.EntityType -join ';')
        }
    } | Sort-Object Name) (Join-Path $OutDir 'tag-categories.csv') @('Name','Description','Cardinality','EntityType')

    Write-Meta ($tagRows | ForEach-Object {
        [pscustomobject]@{ Category = $_.Category.Name; Name = $_.Name; Description = $_.Description }
    } | Sort-Object Category,Name) (Join-Path $OutDir 'tags.csv') @('Category','Name','Description')

    Write-Meta ($arows | Sort-Object EntityType,EntityName,Category,Tag) (Join-Path $OutDir 'tag-assignments.csv') @('EntityType','EntityName','EntityUuid','Datacenter','FolderType','EntityPath','Category','Tag','MoRef')
}

# --- 4. Custom attributes ---
if ($Include -contains 'CustomAttributes') {
    $cfm = Get-View (Get-View ServiceInstance -Server $vc).Content.CustomFieldsManager -Server $vc
    $fieldByKey = @{}
    $defs = foreach ($f in $cfm.Field) {
        $fieldByKey[$f.Key] = $f.Name
        [pscustomobject]@{
            Name       = $f.Name
            TargetType = if ([string]::IsNullOrEmpty($f.ManagedObjectType)) { 'Global' } else { $f.ManagedObjectType }
            Key        = $f.Key
        }
    }
    $types = 'VirtualMachine','HostSystem','Datastore','ClusterComputeResource','Datacenter','Folder','ResourcePool','StoragePod','DistributedVirtualPortgroup'
    $vrows = New-Object System.Collections.ArrayList
    foreach ($t in $types) {
        $views = @()
        try { $views = Get-View -ViewType $t -Property Name,CustomValue -Server $vc } catch { continue }
        $uuidMap = @{}
        if ($t -eq 'VirtualMachine') {
            foreach ($u in (Get-View -ViewType VirtualMachine -Property Config.InstanceUuid -Server $vc)) {
                $uuidMap[$u.MoRef.ToString()] = $u.Config.InstanceUuid
            }
        }
        foreach ($v in $views) {
            if (-not $v.CustomValue) { continue }
            $id = $v.MoRef.ToString()
            if ($scopeFolderIds -and -not ($scopeFolderIds.ContainsKey($id) -or $scopeVmIds.ContainsKey($id))) { continue }
            $fi = if ($folderInfo.ContainsKey($id)) { $folderInfo[$id] } else { $null }
            $uuid = if ($uuidMap.ContainsKey($id)) { $uuidMap[$id] } else { '' }
            foreach ($cv in $v.CustomValue) {
                if ([string]::IsNullOrWhiteSpace($cv.Value)) { continue }
                [void]$vrows.Add([pscustomobject]@{
                    EntityType    = $t
                    EntityName    = $v.Name
                    EntityUuid    = $uuid
                    Datacenter    = if ($fi) { $fi.dc }   else { '' }
                    FolderType    = if ($fi) { $fi.type } else { '' }
                    EntityPath    = if ($fi) { $fi.path } else { '' }
                    AttributeName = $fieldByKey[$cv.Key]
                    Value         = $cv.Value
                    MoRef         = $id
                })
            }
        }
    }
    # Likewise, with -Folder keep only attribute definitions that are used in scope
    $usedAttr = @{}
    foreach ($r in $vrows) { $usedAttr[$r.AttributeName] = $true }
    $defRows = $defs | Where-Object { -not $scopeFolderIds -or $AllDefinitions -or $usedAttr.ContainsKey($_.Name) }
    Write-Meta ($defRows | Sort-Object TargetType,Name) (Join-Path $OutDir 'custom-attributes.csv') @('Name','TargetType','Key')

    Write-Meta ($vrows | Sort-Object EntityType,EntityName,AttributeName) (Join-Path $OutDir 'custom-attribute-values.csv') @('EntityType','EntityName','EntityUuid','Datacenter','FolderType','EntityPath','AttributeName','Value','MoRef')
}

Disconnect-VIServer -Server $vc -Confirm:$false | Out-Null
Write-Host ""
Write-Host "[+] Export complete -> $OutDir"
