<#
.SYNOPSIS
  把 Export-VcMeta.ps1 產生的 CSV 匯入到目標 vCenter：
  重建 Folder 樹、Tag 分類/標籤/指派、Custom Attribute 定義/值、VM Notes，
  並可選擇把 VM 搬到對應的 Folder。

.EXAMPLE
  # 先試跑（不改任何東西）
  ./Import-VcMeta.ps1 -Server 10.0.1.19 -User administrator@vsphere.local -Password '<target-pw>' -InDir .\export-A -DryRun

  # 正式套用，含把 VM 搬進資料夾
  ./Import-VcMeta.ps1 -Server 10.0.1.19 -User administrator@vsphere.local -Password '<target-pw>' -InDir .\export-A -MoveVMs

.NOTES
  比對規則：VM 先用 InstanceUuid，找不到再用名稱；其他物件用名稱；Folder 用 Datacenter+類型+路徑。
  Datacenter 名稱不同時用 -DatacenterMap 'OldDC=NewDC'。
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$Server,
    [string]$User,
    [string]$Password,
    [pscredential]$Credential,
    [Parameter(Mandatory)][string]$InDir,
    [ValidateSet('Folders','VMPlacement','Tags','CustomAttributes','Notes')]
    [string[]]$Include = @('Folders','Tags','CustomAttributes','Notes'),
    [string[]]$DatacenterMap,          # 例：-DatacenterMap 'DC-A=DC-B','Lab=Lab2'
    [string[]]$Folder,                 # 只匯入這些 VM 資料夾(含子樹)與裡面的 VM，例：-Folder 'Linux'
    [string[]]$OnlyVMs,                # 只處理這些名字的 VM、其他型別全跳過；Register/Move 註冊完補資料用
    [switch]$MoveVMs,                  # 把 VM 搬進對應 Folder（預設不搬）
    [switch]$DryRun,
    [string]$ReportPath
)

$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding $false } catch { }
$enc = if ($PSVersionTable.PSVersion.Major -ge 6) { 'utf8BOM' } else { 'UTF8' }
$InDir = (Resolve-Path $InDir).Path
if (-not $ReportPath) { $ReportPath = Join-Path $InDir ('import-report-{0:yyyyMMdd-HHmmss}.csv' -f (Get-Date)) }

$script:Report = New-Object System.Collections.ArrayList
$script:Stats  = @{}
function Add-Result {
    param([string]$Section,[string]$Action,[string]$Target,[string]$Detail = '')
    [void]$script:Report.Add([pscustomobject]@{
        Section = $Section; Action = $Action; Target = $Target; Detail = $Detail
    })
    $k = "$Section/$Action"
    if (-not $script:Stats.ContainsKey($k)) { $script:Stats[$k] = 0 }
    $script:Stats[$k]++
}
function Read-Meta {
    param([string]$Name)
    $p = Join-Path $InDir $Name
    if (-not (Test-Path $p)) { Write-Host "[!] 缺少 $Name，略過"; return @() }
    @(Import-Csv -Path $p)
}
# --- 資料夾範圍（-Folder）---
function Test-PathInScope {
    param([string]$Path)
    if (-not $Folder) { return $true }
    foreach ($f in $Folder) {
        $fn = $f.Trim('/')
        if ($Path -eq $fn -or $Path -like "$fn/*") { return $true }
    }
    return $false
}
$script:ScopeVmName = @{}
$script:ScopeVmUuid = @{}
if ($Folder) {
    foreach ($r in (Read-Meta 'vm-placement.csv')) {
        if (-not (Test-PathInScope $r.FolderPath)) { continue }
        $script:ScopeVmName[$r.VMName] = $true
        if ($r.InstanceUuid) { $script:ScopeVmUuid[$r.InstanceUuid] = $true }
    }
}
function Test-VmInScope {
    param([string]$Name, [string]$Uuid)
    if ($OnlyVMs -and ($OnlyVMs -notcontains $Name)) { return $false }
    if (-not $Folder) { return $true }
    if ($Uuid -and $script:ScopeVmUuid.ContainsKey($Uuid)) { return $true }
    return $script:ScopeVmName.ContainsKey($Name)
}
# tag/屬性的指派：資料夾看路徑、VM 看是否在範圍內，其他型別在 -Folder 模式下一律跳過
# （-OnlyVMs：只處理這些 VM，其他型別一律跳過 —— Register/Move 註冊完補資料用）
function Test-EntityInScope {
    param([string]$EntityType, [string]$EntityName, [string]$EntityUuid, [string]$EntityPath)
    if ($OnlyVMs -and $EntityType -ne 'VirtualMachine') { return $false }
    if ($EntityType -eq 'VirtualMachine' -and -not (Test-VmInScope $EntityName $EntityUuid)) { return $false }
    if (-not $Folder) { return $true }
    switch ($EntityType) {
        'Folder'         { return (Test-PathInScope $EntityPath) }
        'VirtualMachine' { return $true }
        default          { return $false }
    }
}

function Map-Dc {
    param([string]$Name)
    foreach ($m in $DatacenterMap) {
        $kv = $m -split '=', 2
        if ($kv.Count -eq 2 -and $kv[0] -eq $Name) { return $kv[1] }
    }
    return $Name
}

# --- 連線 ---
Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -DisplayDeprecationWarnings $false -Scope Session -Confirm:$false | Out-Null
if (-not $Credential -and $User) {
    $sec = ConvertTo-SecureString $Password -AsPlainText -Force
    $Credential = New-Object System.Management.Automation.PSCredential($User, $sec)
}
Write-Host "[*] 連線 $Server ..."
$vc = Connect-VIServer -Server $Server -Credential $Credential
Write-Host "[+] 已連線 $($vc.Name)  ($($vc.Version) build $($vc.Build))"
if ($DryRun) { Write-Host "[!] DryRun 模式：只比對、不寫入" -ForegroundColor Yellow }

# --- 目標端查詢表（用到才載入） ---
$script:LK = @{}
function Get-LK {
    param([string]$Kind)
    if ($script:LK.ContainsKey($Kind)) { return $script:LK[$Kind] }
    Write-Host "    (載入目標端 $Kind 清單...)"
    $tbl = @{}
    switch ($Kind) {
        'VMByUuid'   { foreach ($v in (Get-View -ViewType VirtualMachine -Property Config.InstanceUuid -Server $vc)) {
                           if ($v.Config.InstanceUuid) { $tbl[$v.Config.InstanceUuid] = $v.MoRef.ToString() } } }
        'VMByName'   { foreach ($v in (Get-VM -Server $vc)) { if (-not $tbl.ContainsKey($v.Name)) { $tbl[$v.Name] = @() }; $tbl[$v.Name] += $v }
                       foreach ($v in (Get-Template -Server $vc)) { if (-not $tbl.ContainsKey($v.Name)) { $tbl[$v.Name] = @() }; $tbl[$v.Name] += $v } }
        'VMHost'     { foreach ($h in (Get-VMHost -Server $vc))      { $tbl[$h.Name] = $h } }
        'Datastore'  { foreach ($d in (Get-Datastore -Server $vc))   { $tbl[$d.Name] = $d } }
        'Cluster'    { foreach ($c in (Get-Cluster -Server $vc))     { $tbl[$c.Name] = $c } }
        'Datacenter' { foreach ($d in (Get-Datacenter -Server $vc))  { $tbl[$d.Name] = $d } }
        'ResourcePool' { foreach ($r in (Get-ResourcePool -Server $vc)) { $tbl[$r.Name] = $r } }
        'DatastoreCluster' { foreach ($s in (Get-DatastoreCluster -Server $vc)) { $tbl[$s.Name] = $s } }
        'Portgroup'  { foreach ($p in (Get-VDPortgroup -Server $vc -ErrorAction SilentlyContinue)) { $tbl[$p.Name] = $p } }
        'VApp'       { foreach ($a in (Get-VApp -Server $vc -ErrorAction SilentlyContinue)) { $tbl[$a.Name] = $a } }
    }
    $script:LK[$Kind] = $tbl
    return $tbl
}

# Folder 快取： key = "<dc>|<type>|<path>"  -> Folder 物件
$script:FolderCache = @{}
$script:WouldFolder = @{}   # DryRun 用：已回報過「會建立」的資料夾路徑
function Get-RootFolder {
    param([string]$DcName,[string]$FolderType)
    $dc = (Get-LK 'Datacenter')[$DcName]
    if (-not $dc) { return $null }
    $moref = switch ($FolderType) {
        'VM'             { $dc.ExtensionData.VmFolder }
        'HostAndCluster' { $dc.ExtensionData.HostFolder }
        'Datastore'      { $dc.ExtensionData.DatastoreFolder }
        'Network'        { $dc.ExtensionData.NetworkFolder }
        default          { $null }
    }
    if (-not $moref) { return $null }
    Get-Folder -Id $moref.ToString() -Server $vc -ErrorAction SilentlyContinue
}
function Resolve-FolderPath {
    param([string]$DcName,[string]$FolderType,[string]$Path,[switch]$Create)
    if ([string]::IsNullOrEmpty($Path)) { return (Get-RootFolder $DcName $FolderType) }
    $key = "$DcName|$FolderType|$Path"
    if ($script:FolderCache.ContainsKey($key)) { return $script:FolderCache[$key] }
    $parent = Get-RootFolder $DcName $FolderType
    if (-not $parent) { return $null }
    $acc = @()
    foreach ($seg in ($Path -split '/')) {
        $acc += $seg
        $subKey = "$DcName|$FolderType|$($acc -join '/')"
        if ($script:FolderCache.ContainsKey($subKey)) { $parent = $script:FolderCache[$subKey]; continue }
        $child = Get-Folder -Location $parent -Name $seg -NoRecursion -Server $vc -ErrorAction SilentlyContinue |
                 Select-Object -First 1
        if (-not $child) {
            if (-not $Create) { return $null }
            if ($DryRun) {
                # 這一段不存在 => 之後每一段也都不存在，一次把剩下的路徑都列出來（不重複列）
                $walk = @()
                foreach ($seg2 in ($Path -split '/')) {
                    $walk += $seg2
                    $k = "$DcName|$FolderType|$($walk -join '/')"
                    if ($walk.Count -lt $acc.Count) { continue }
                    if (-not $script:WouldFolder.ContainsKey($k)) {
                        $script:WouldFolder[$k] = $true
                        Add-Result 'Folder' 'WouldCreate' $k
                    }
                }
                return $null
            }
            $child = New-Folder -Name $seg -Location $parent -Server $vc
            Add-Result 'Folder' 'Created' $subKey
        }
        $script:FolderCache[$subKey] = $child
        $parent = $child
    }
    return $parent
}

function Resolve-Entity {
    param([string]$Type,[string]$Name,[string]$Uuid,[string]$DcName,[string]$FolderType,[string]$Path)
    switch ($Type) {
        'VirtualMachine' {
            if ($Uuid) {
                $id = (Get-LK 'VMByUuid')[$Uuid]
                if ($id) {
                    # 用 MoRef 轉物件：Get-VM -Id 對 Template 會失敗
                    try {
                        $mo = New-Object VMware.Vim.ManagedObjectReference($id)
                        $obj = Get-VIObjectByVIView -MORef $mo -Server $vc -ErrorAction SilentlyContinue
                        if ($obj) { return $obj }
                    } catch { }
                    $byId = Get-VM -Id $id -Server $vc -ErrorAction SilentlyContinue
                    if ($byId) { return $byId }
                }
            }
            $hit = (Get-LK 'VMByName')[$Name]
            if ($hit -and @($hit).Count -eq 1) { return @($hit)[0] }
            if ($hit) { return $null }   # 同名多台 -> 視為無法判定
            return $null
        }
        'HostSystem'                 { return (Get-LK 'VMHost')[$Name] }
        'VMHost'                     { return (Get-LK 'VMHost')[$Name] }
        'Datastore'                  { return (Get-LK 'Datastore')[$Name] }
        'ClusterComputeResource'     { return (Get-LK 'Cluster')[$Name] }
        'Cluster'                    { return (Get-LK 'Cluster')[$Name] }
        'Datacenter'                 { return (Get-LK 'Datacenter')[(Map-Dc $Name)] }
        'ResourcePool'               { return (Get-LK 'ResourcePool')[$Name] }
        'StoragePod'                 { return (Get-LK 'DatastoreCluster')[$Name] }
        'DatastoreCluster'           { return (Get-LK 'DatastoreCluster')[$Name] }
        'DistributedVirtualPortgroup'{ return (Get-LK 'Portgroup')[$Name] }
        'VirtualApp'                 { return (Get-LK 'VApp')[$Name] }
        'VApp'                       { return (Get-LK 'VApp')[$Name] }
        'Folder'                     { return (Resolve-FolderPath (Map-Dc $DcName) $FolderType $Path) }
        default                      { return $null }
    }
}

# =====================================================================
# 1. Folders
# =====================================================================
if ($Include -contains 'Folders') {
    Write-Host "`n[1] Folder 結構"
    $rows = Read-Meta 'folders.csv' |
            Where-Object { -not $Folder -or ($_.FolderType -eq 'VM' -and (Test-PathInScope $_.Path)) } |
            Sort-Object { [int]$_.Depth }, Path
    foreach ($r in $rows) {
        $dcName = Map-Dc $r.Datacenter
        if (-not (Get-LK 'Datacenter')[$dcName]) {
            Add-Result 'Folder' 'NoDatacenter' "$dcName|$($r.FolderType)|$($r.Path)"
            continue
        }
        $existing = Resolve-FolderPath $dcName $r.FolderType $r.Path
        if ($existing) { Add-Result 'Folder' 'Exists' "$dcName|$($r.FolderType)|$($r.Path)"; continue }
        $null = Resolve-FolderPath $dcName $r.FolderType $r.Path -Create
    }
}

# =====================================================================
# 2. Custom Attributes（定義 + 值）
# =====================================================================
if ($Include -contains 'CustomAttributes') {
    Write-Host "`n[2] Custom Attributes"
    $targetTypeMap = @{
        'VirtualMachine'         = 'VirtualMachine'
        'HostSystem'             = 'VMHost'
        'ClusterComputeResource' = 'Cluster'
        'Datastore'              = 'Datastore'
        'Datacenter'             = 'Datacenter'
        'Folder'                 = 'Folder'
        'ResourcePool'           = 'ResourcePool'
        'StoragePod'             = 'DatastoreCluster'
        'VirtualApp'             = 'VApp'
    }
    $existingCA = @{}
    foreach ($c in (Get-CustomAttribute -Server $vc)) { $existingCA["$($c.Name)|$($c.TargetType)"] = $c; $existingCA[$c.Name] = $c }

    # 先算出範圍內的屬性值，-Folder 模式下只建「這些值用得到」的屬性定義
    $caValRows = @(Read-Meta 'custom-attribute-values.csv' |
                   Where-Object { Test-EntityInScope $_.EntityType $_.EntityName $_.EntityUuid $_.EntityPath })
    $usedAttr = @{}
    foreach ($r in $caValRows) { $usedAttr[$r.AttributeName] = $true }

    foreach ($d in (Read-Meta 'custom-attributes.csv')) {
        if (($Folder -or $OnlyVMs) -and -not $usedAttr.ContainsKey($d.Name)) { continue }
        if ($existingCA.ContainsKey($d.Name)) { Add-Result 'CustomAttribute' 'Exists' $d.Name; continue }
        if ($DryRun) { Add-Result 'CustomAttribute' 'WouldCreate' $d.Name $d.TargetType; continue }
        try {
            if ($d.TargetType -eq 'Global' -or -not $targetTypeMap.ContainsKey($d.TargetType)) {
                $new = New-CustomAttribute -Name $d.Name -Server $vc
            } else {
                $new = New-CustomAttribute -Name $d.Name -TargetType $targetTypeMap[$d.TargetType] -Server $vc
            }
            $existingCA[$d.Name] = $new
            Add-Result 'CustomAttribute' 'Created' $d.Name $d.TargetType
        } catch {
            Add-Result 'CustomAttribute' 'Failed' $d.Name $_.Exception.Message
        }
    }

    foreach ($v in $caValRows) {
        $e = Resolve-Entity $v.EntityType $v.EntityName $v.EntityUuid $v.Datacenter $v.FolderType $v.EntityPath
        if (-not $e) { Add-Result 'CAValue' 'EntityNotFound' "$($v.EntityType):$($v.EntityName)" $v.AttributeName; continue }
        $cur = ($e | Get-Annotation -CustomAttribute $v.AttributeName -Server $vc -ErrorAction SilentlyContinue).Value
        if ($cur -eq $v.Value) { Add-Result 'CAValue' 'AlreadySet' "$($v.EntityType):$($v.EntityName)" $v.AttributeName; continue }
        if ($DryRun) { Add-Result 'CAValue' 'WouldSet' "$($v.EntityType):$($v.EntityName)" "$($v.AttributeName)=$($v.Value)"; continue }
        try {
            $null = $e | Set-Annotation -CustomAttribute $v.AttributeName -Value $v.Value -Server $vc
            Add-Result 'CAValue' 'Set' "$($v.EntityType):$($v.EntityName)" "$($v.AttributeName)=$($v.Value)"
        } catch {
            Add-Result 'CAValue' 'Failed' "$($v.EntityType):$($v.EntityName)" $_.Exception.Message
        }
    }
}

# =====================================================================
# 3. Tags（分類 / 標籤 / 指派）
# =====================================================================
if ($Include -contains 'Tags') {
    Write-Host "`n[3] Tags"
    $wouldCat = @{}; $wouldTag = @{}   # DryRun 用：假設會被建立的分類/標籤
    # 先算出範圍內的指派，-Folder 模式下只建「這些指派用得到」的分類與標籤
    $assignRows = @(Read-Meta 'tag-assignments.csv' |
                    Where-Object { Test-EntityInScope $_.EntityType $_.EntityName $_.EntityUuid $_.EntityPath })
    $usedCat = @{}; $usedTagKey = @{}
    foreach ($r in $assignRows) { $usedCat[$r.Category] = $true; $usedTagKey["$($r.Category)|$($r.Tag)"] = $true }

    $catByName = @{}
    foreach ($c in (Get-TagCategory -Server $vc)) { $catByName[$c.Name] = $c }
    foreach ($c in (Read-Meta 'tag-categories.csv')) {
        if (($Folder -or $OnlyVMs) -and -not $usedCat.ContainsKey($c.Name)) { continue }
        if ($catByName.ContainsKey($c.Name)) { Add-Result 'TagCategory' 'Exists' $c.Name; continue }
        if ($DryRun) { $wouldCat[$c.Name] = $true; Add-Result 'TagCategory' 'WouldCreate' $c.Name $c.EntityType; continue }
        try {
            $et = if ([string]::IsNullOrWhiteSpace($c.EntityType)) { @('All') } else { $c.EntityType -split ';' }
            $new = New-TagCategory -Name $c.Name -Description $c.Description -Cardinality $c.Cardinality -EntityType $et -Server $vc
            $catByName[$c.Name] = $new
            Add-Result 'TagCategory' 'Created' $c.Name $c.EntityType
        } catch {
            Add-Result 'TagCategory' 'Failed' $c.Name $_.Exception.Message
        }
    }

    $tagByKey = @{}
    foreach ($t in (Get-Tag -Server $vc)) { $tagByKey["$($t.Category.Name)|$($t.Name)"] = $t }
    foreach ($t in (Read-Meta 'tags.csv')) {
        $key = "$($t.Category)|$($t.Name)"
        if (($Folder -or $OnlyVMs) -and -not $usedTagKey.ContainsKey($key)) { continue }
        if ($tagByKey.ContainsKey($key)) { Add-Result 'Tag' 'Exists' $key; continue }
        if (-not $catByName.ContainsKey($t.Category) -and -not $wouldCat.ContainsKey($t.Category)) { Add-Result 'Tag' 'NoCategory' $key; continue }
        if ($DryRun) { $wouldTag[$key] = $true; Add-Result 'Tag' 'WouldCreate' $key; continue }
        try {
            $new = New-Tag -Name $t.Name -Category $catByName[$t.Category] -Description $t.Description -Server $vc
            $tagByKey[$key] = $new
            Add-Result 'Tag' 'Created' $key
        } catch {
            Add-Result 'Tag' 'Failed' $key $_.Exception.Message
        }
    }

    # $assignRows 已在前面依 -Folder 範圍過濾好
    if ($assignRows.Count) {
        # 先嘗試一次抓完既有指派；環境若有無法存取的物件(例如死掉的 datastore)會整個中斷，
        # 那就改成逐一比對。
        $existingAssign = @{}
        $bulkOk = $false
        try {
            foreach ($a in (Get-TagAssignment -Server $vc)) {
                $id = ''
                try { $id = $a.Entity.ExtensionData.MoRef.ToString() } catch { }
                $existingAssign["$id|$($a.Tag.Category.Name)|$($a.Tag.Name)"] = $true
            }
            $bulkOk = $true
        } catch {
            Write-Host "    [!] 無法一次列出既有 tag 指派($($_.Exception.Message.Split([char]10)[0]))，改逐一比對"
        }
        foreach ($a in $assignRows) {
            $key = "$($a.Category)|$($a.Tag)"
            if (-not $tagByKey.ContainsKey($key) -and -not $wouldTag.ContainsKey($key)) { Add-Result 'TagAssignment' 'NoTag' "$($a.EntityName) <- $key"; continue }
            $e = Resolve-Entity $a.EntityType $a.EntityName $a.EntityUuid $a.Datacenter $a.FolderType $a.EntityPath
            if (-not $e) { Add-Result 'TagAssignment' 'EntityNotFound' "$($a.EntityType):$($a.EntityName)" $key; continue }
            $eid = ''
            try { $eid = $e.ExtensionData.MoRef.ToString() } catch { }
            if ($bulkOk) {
                if ($existingAssign.ContainsKey("$eid|$key")) { Add-Result 'TagAssignment' 'Exists' "$($a.EntityName)" $key; continue }
            } else {
                $already = $null
                try { $already = Get-TagAssignment -Entity $e -Tag $tagByKey[$key] -Server $vc -ErrorAction SilentlyContinue } catch { }
                if ($already) { Add-Result 'TagAssignment' 'Exists' "$($a.EntityName)" $key; continue }
            }
            if ($DryRun) { Add-Result 'TagAssignment' 'WouldAssign' "$($a.EntityType):$($a.EntityName)" $key; continue }
            try {
                $null = New-TagAssignment -Tag $tagByKey[$key] -Entity $e -Server $vc -Confirm:$false
                Add-Result 'TagAssignment' 'Assigned' "$($a.EntityType):$($a.EntityName)" $key
            } catch {
                Add-Result 'TagAssignment' 'Failed' "$($a.EntityType):$($a.EntityName)" $_.Exception.Message
            }
        }
    }
}

# =====================================================================
# 4. VM Notes
# =====================================================================
if ($Include -contains 'Notes') {
    Write-Host "`n[4] VM Notes"
    foreach ($n in (Read-Meta 'notes.csv' | Where-Object { Test-VmInScope $_.EntityName $_.InstanceUuid })) {
        $vm = Resolve-Entity 'VirtualMachine' $n.EntityName $n.InstanceUuid '' '' ''
        if (-not $vm) { Add-Result 'Notes' 'VMNotFound' $n.EntityName; continue }
        $curNotes = $null
        try { $curNotes = $vm.ExtensionData.Config.Annotation } catch { }
        if ($null -eq $curNotes) { $curNotes = $vm.Notes }
        if ($curNotes -eq $n.Notes) { Add-Result 'Notes' 'AlreadySet' $n.EntityName; continue }
        if ($DryRun) { Add-Result 'Notes' 'WouldSet' $n.EntityName ($n.Notes.Substring(0,[Math]::Min(60,$n.Notes.Length))); continue }
        try {
            if ($vm.GetType().Name -match 'Template') {
                # Template 沒有 Set-VM -Notes，改用 ReconfigVM 設 annotation
                $spec = New-Object VMware.Vim.VirtualMachineConfigSpec
                $spec.Annotation = $n.Notes
                $vm.ExtensionData.ReconfigVM($spec)
            } else {
                $null = Set-VM -VM $vm -Notes $n.Notes -Server $vc -Confirm:$false
            }
            Add-Result 'Notes' 'Set' $n.EntityName
        } catch {
            Add-Result 'Notes' 'Failed' $n.EntityName $_.Exception.Message
        }
    }
}

# =====================================================================
# 5. VM 位置（搬進 Folder）
# =====================================================================
if (($Include -contains 'VMPlacement') -or $MoveVMs) {
    Write-Host "`n[5] VM 位置"
    if (-not $MoveVMs) {
        Write-Host "    (未指定 -MoveVMs，只比對不搬移)"
    }
    foreach ($p in (Read-Meta 'vm-placement.csv' | Where-Object { (-not $Folder -or (Test-PathInScope $_.FolderPath)) -and (-not $OnlyVMs -or $OnlyVMs -contains $_.VMName) })) {
        if ([string]::IsNullOrEmpty($p.FolderPath)) { continue }
        $vm = Resolve-Entity 'VirtualMachine' $p.VMName $p.InstanceUuid '' '' ''
        if (-not $vm) { Add-Result 'VMPlacement' 'VMNotFound' $p.VMName; continue }
        $dcName = Map-Dc $p.Datacenter
        $target = Resolve-FolderPath $dcName 'VM' $p.FolderPath
        if (-not $target) { Add-Result 'VMPlacement' 'FolderNotFound' $p.VMName "$dcName|$($p.FolderPath)"; continue }
        $curFolderId = ''
        try { $curFolderId = $vm.ExtensionData.Parent.ToString() } catch { }
        if ($curFolderId -eq $target.ExtensionData.MoRef.ToString()) { Add-Result 'VMPlacement' 'AlreadyThere' $p.VMName; continue }
        if ($DryRun -or -not $MoveVMs) { Add-Result 'VMPlacement' 'WouldMove' $p.VMName $p.FolderPath; continue }
        try {
            if ($vm.GetType().Name -match 'Template') {
                $null = Move-Template -Template $vm -Destination $target -Server $vc -Confirm:$false
            } else {
                $null = Move-VM -VM $vm -InventoryLocation $target -Server $vc -Confirm:$false
            }
            Add-Result 'VMPlacement' 'Moved' $p.VMName $p.FolderPath
        } catch {
            Add-Result 'VMPlacement' 'Failed' $p.VMName $_.Exception.Message
        }
    }
}

# --- 總結 ---
Write-Host "`n================ 結果 ================"
foreach ($k in ($script:Stats.Keys | Sort-Object)) {
    Write-Host ("  {0,-34} {1,5}" -f $k, $script:Stats[$k])
}
if ($script:Report.Count) {
    $script:Report | Export-Csv -Path $ReportPath -NoTypeInformation -Encoding $enc
    Write-Host "`n[+] 明細報告 -> $ReportPath"
}
Disconnect-VIServer -Server $vc -Confirm:$false | Out-Null
