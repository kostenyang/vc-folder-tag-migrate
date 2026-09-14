<#
.SYNOPSIS
  把 VM / 範本註冊回新 vCenter 的 inventory，並直接放進原資料夾。兩種模式：

  模式 B（主線）-UnregisteredCsv：依 Unregister-VmFromOldVc 產生的 unregistered.csv 逐台註冊。
    精確只註冊「舊 vC 拔掉的那批」，vmx 路徑、名稱、原資料夾都從 CSV 來；CSV 可先用 Excel 改再跑。
  模式 A -Datastore：掃 datastore 上所有 .vmx / .vmtx，還沒在 inventory 的都註冊（孤兒清理、沒有 CSV 時用）。

  結尾對帳：CSV（或匯出清單）裡還沒出現在新 vC 的 VM 列成 PendingOnSource。
  已註冊的依 [datastore] 路徑比對一律跳過，重跑安全；.vmtx 註冊成範本；只註冊、不開機。

.EXAMPLE
  # 主線：先看
  ./Register-VmxFromDatastore.ps1 -Server <新vC> -Password '<pw>' -Cluster cl01 -UnregisteredCsv .\export-A\unregistered.csv -DryRun

  # 主線：正式（註冊 + 放進原資料夾）
  ./Register-VmxFromDatastore.ps1 -Server <新vC> -Password '<pw>' -Cluster cl01 -UnregisteredCsv .\export-A\unregistered.csv

  # 掃 datastore（沒有 CSV 時）
  ./Register-VmxFromDatastore.ps1 -Server <新vC> -Password '<pw>' -Datastore ds01 -Cluster cl01 -PlacementCsv .\export-A\vm-placement.csv -DryRun

  # 可選：把「補屬性 / Notes / tag」併進來一次做（主線是分開跑 Import-VcMeta）
  ./Register-VmxFromDatastore.ps1 ... -UnregisteredCsv .\export-A\unregistered.csv -MetaDir .\export-A -DatacenterMap 'DC-A=DC-B'
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$Server,
    [string]$User,
    [string]$Password,
    [pscredential]$Credential,

    [string[]]$Datastore,              # 模式 A：掃這些 datastore 上的 vmx/vmtx
    [string]$UnregisteredCsv,          # 模式 B（主線）：依 Unregister-VmFromOldVc 產生的 unregistered.csv 逐台註冊，直接放進原資料夾
    [string]$Cluster,                  # 註冊到這個叢集（挑有掛該 datastore 的主機，輪流用）
    [string]$VMHost,                   # 或指定單一主機
    [string]$ResourcePool,             # 預設用主機所屬叢集 / 主機的根 resource pool

    [string]$Folder,                   # 註冊到哪個 VM 資料夾（相對 datacenter 的路徑，例 'Linux'），預設 DC 根
    [string]$PlacementCsv,             # Export-VcMeta 的 vm-placement.csv：註冊後依 VMName 搬到各自資料夾
    [switch]$CreateFolders,            # 資料夾不存在就建

    # --- 照舊 vC 把資料夾 / 自訂屬性 / Notes / tag 一起做好（二選一）---
    [string]$SourceServer,             # 直接連舊 vC 抓（會先跑 Export-VcMeta 到 -MetaDir）；VM 必須還在舊 vC inventory 裡才抓得到
    [string]$SourceUser = 'administrator@vsphere.local',
    [string]$SourcePassword,
    [string]$MetaDir,                  # 或給已經匯出好的目錄（Export-VcMeta 的 -OutDir）—— 建議走這條：先匯出再 unregister
    [string[]]$DatacenterMap,          # 舊=新 DC 名稱對應，例 'DC-A=DC-B'
    [string[]]$SourceDatacenter,       # 只從舊 vC 抓這些 Datacenter

    [string[]]$Include,                # 只註冊符合的名稱（wildcard，例 'web*','db01'）
    [string[]]$Exclude = @('vCLS*'),   # 排除（預設略過 vCLS）
    [switch]$NameFromFile,             # 用 vmx 檔名當 VM 名稱（預設抓 vmx 裡的 displayName，讀不到才用檔名）
    [switch]$NoTemplates,
    [switch]$DryRun,
    [string]$ReportPath
)

$ErrorActionPreference = 'Stop'
if (-not $Datastore -and -not $UnregisteredCsv) { throw "要給 -UnregisteredCsv（主線）或 -Datastore（掃 datastore）其中一個" }
try { [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding $false } catch { }
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
    Write-Host ("  {0} {1,-14} {2}  {3}" -f $tag, $Action, $Target, $Detail)
}
function Test-NameMatch {
    param([string]$Name)
    if ($Include) { $hit = $false; foreach ($p in $Include) { if ($Name -like $p) { $hit = $true; break } }; if (-not $hit) { return $false } }
    foreach ($p in $Exclude) { if ($Name -like $p) { return $false } }
    return $true
}

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$script:RegisteredNames = New-Object System.Collections.ArrayList

# --- 照舊 vC：先把舊 vC 的資料抓下來（或用現成的 -MetaDir）---
if ($SourceServer) {
    if (-not $SourcePassword) { throw "-SourceServer 需要 -SourcePassword" }
    if (-not $MetaDir) { $MetaDir = Join-Path (Get-Location) ('meta-{0}-{1:yyyyMMdd-HHmmss}' -f ($SourceServer -replace '[^\w.-]','_'), (Get-Date)) }
    Write-Host "==================== 從舊 vC $SourceServer 抓資料夾 / tag / 屬性 / Notes ===================="
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

# --- 連線 ---
Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -DisplayDeprecationWarnings $false -Scope Session -Confirm:$false | Out-Null
if (-not $Credential -and $User) {
    $sec = ConvertTo-SecureString $Password -AsPlainText -Force
    $Credential = New-Object System.Management.Automation.PSCredential($User, $sec)
}
Write-Host "[*] 連線 $Server ..."
$vc = Connect-VIServer -Server $Server -Credential $Credential
Write-Host "[+] 已連線 $($vc.Name)  ($($vc.Version) build $($vc.Build))"
if ($DryRun) { Write-Host "[!] DryRun 模式：只掃描、不註冊" -ForegroundColor Yellow }

# --- 已註冊的 vmx 路徑（跳過用）---
$registered = @{}
foreach ($v in (Get-View -ViewType VirtualMachine -Property Config.Files.VmPathName -Server $vc)) {
    $p = $v.Config.Files.VmPathName
    if ($p) { $registered[$p.ToLower()] = $true }
}
Write-Host "[*] inventory 目前有 $($registered.Count) 台 VM/範本"

# --- 資料夾解析（同 Import-VcMeta 的規則）---
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

# --- 讀 vmx 裡的 displayName（RegisterVM 的 name 不能空，傳 $null 會變 '' 被拒）---
function Get-VmxDisplayName {
    param([string]$DcName, [string]$DsName, [string]$RelPath)   # RelPath 例：dir/vm.vmx
    $tmp = Join-Path ([IO.Path]::GetTempPath()) ("vmx-" + [guid]::NewGuid().ToString('N') + ".vmx")
    try {
        Copy-DatastoreItem -Item ("vmstore:\{0}\{1}\{2}" -f $DcName, $DsName, ($RelPath -replace '/', '\')) -Destination $tmp -Force -ErrorAction Stop | Out-Null
        $m = Select-String -Path $tmp -Pattern '^\s*displayName\s*=\s*"(.*)"\s*$' | Select-Object -First 1
        if ($m) { return $m.Matches[0].Groups[1].Value }
    } catch { }
    finally { Remove-Item $tmp -Force -ErrorAction SilentlyContinue }
    return $null
}

# --- placement 對照（VMName -> FolderPath）---
$placement = @{}
if ($PlacementCsv) {
    foreach ($r in (Import-Csv $PlacementCsv)) {
        if ($r.FolderPath) { $placement[$r.VMName] = $r.FolderPath }
    }
    Write-Host "[*] placement 對照 $($placement.Count) 筆（$PlacementCsv）"
}

# =====================================================================
# 模式 B：依 unregistered.csv 逐台註冊（主線）
# =====================================================================
$csvRows = @()
if ($UnregisteredCsv) {
    $UnregisteredCsv = (Resolve-Path $UnregisteredCsv).Path
    $csvRows = @(Import-Csv $UnregisteredCsv)
    if (-not $csvRows.Count) { throw "$UnregisteredCsv 是空的" }
    foreach ($col in 'VMName','VmPathName') { if ($csvRows[0].PSObject.Properties.Name -notcontains $col) { throw "$UnregisteredCsv 缺 $col 欄位" } }
    $hasFolder = $csvRows[0].PSObject.Properties.Name -contains 'FolderPath'
    $CreateFolders = $true
    Write-Host "`n=== 依 $(Split-Path $UnregisteredCsv -Leaf) 註冊 $($csvRows.Count) 台 ==="
    $dsCache = @{}; $hostCache = @{}; $poolCache = @{}; $rr = @{}
    # 對帳用：這份 CSV 涵蓋到的 datastore
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
                    if (-not $rp) { throw "找不到 resource pool '$ResourcePool'（在 $($h.Parent.Name) 下）" }
                    $rp.ExtensionData.MoRef
                } else { (Get-View $h.ExtensionData.Parent -Property ResourcePool -Server $vc).ResourcePool }
            }
            if ($hs.Count) { Write-Host "  datastore $dsName → 主機: $($hs.Name -join ', ')" }
        }
        $dsObj = $dsCache[$dsName]; $hs = $hostCache[$dsName]
        if (-not $dsObj) { Add-Result 'DatastoreNotFound' $vmx "新 vC 沒有 datastore '$dsName'（還沒掛？）"; continue }
        if (-not $hs.Count) { Add-Result 'NoHost' $vmx "沒有可用主機掛著 $dsName（Cluster=$Cluster VMHost=$VMHost）"; continue }
        $h = $hs[$rr[$dsName] % $hs.Count]; $rr[$dsName]++

        $dc = $dsObj.Datacenter
        $fp = if ($hasFolder) { $row.FolderPath } else { $Folder }
        $dest = Resolve-VmFolder $dc $fp -Create:$CreateFolders
        $kind = if ($isTpl) { 'Template' } else { 'VM' }
        if ($DryRun) {
            [void]$script:RegisteredNames.Add($row.VMName)
            Add-Result 'WouldRegister' $vmx "$kind '$($row.VMName)' -> host=$($h.Name) folder=$(if ($fp) { $fp } else { '(DC 根)' })"
            continue
        }
        if (-not $dest) { Add-Result 'FolderNotFound' $vmx $fp; continue }
        try {
            $fv = Get-View $dest.ExtensionData.MoRef -Server $vc
            $pool = if ($isTpl) { $null } else { $poolCache[$h.Name] }
            $vmRef = $fv.RegisterVM($vmx, $row.VMName, $isTpl, $pool, $h.ExtensionData.MoRef)   # 直接註冊進目的資料夾
            $registered[$vmx.ToLower()] = $true
            [void]$script:RegisteredNames.Add($row.VMName)
            Add-Result 'Registered' $vmx "$kind '$($row.VMName)' on $($h.Name) -> $(if ($fp) { $fp } else { '(DC 根)' })"
        } catch {
            Add-Result 'Failed' $vmx ($_.Exception.Message -split "`n")[0]
        }
    }
    if (-not $PlacementCsv) { $PlacementCsv = $UnregisteredCsv }   # 對帳用
}

# =====================================================================
# 模式 A：掃 datastore
# =====================================================================
foreach ($dsName in ($(if ($UnregisteredCsv) { @() } else { $Datastore }))) {
    Write-Host "`n=== Datastore: $dsName ==="
    $ds = Get-Datastore -Name $dsName -Server $vc
    $dc = $ds.Datacenter

    # 可用主機：有掛這個 datastore、Connected，再依 -Cluster / -VMHost 篩
    $hosts = @($ds | Get-VMHost -Server $vc | Where-Object { $_.ConnectionState -eq 'Connected' })
    if ($Cluster) { $hosts = @($hosts | Where-Object { $_.Parent.Name -eq $Cluster }) }
    if ($VMHost)  { $hosts = @($hosts | Where-Object { $_.Name -eq $VMHost }) }
    if (-not $hosts) { Write-Host "  [!] 沒有可用主機掛著 $dsName（Cluster=$Cluster VMHost=$VMHost），略過"; Add-Result 'NoHost' $dsName; continue }
    Write-Host "  主機: $($hosts.Name -join ', ')"

    # resource pool
    $poolByHost = @{}
    foreach ($h in $hosts) {
        if ($ResourcePool) {
            $rp = Get-ResourcePool -Name $ResourcePool -Location $h.Parent -Server $vc -ErrorAction SilentlyContinue | Select-Object -First 1
            if (-not $rp) { throw "找不到 resource pool '$ResourcePool'（在 $($h.Parent.Name) 下）" }
            $poolByHost[$h.Name] = $rp.ExtensionData.MoRef
        } else {
            $poolByHost[$h.Name] = (Get-View $h.ExtensionData.Parent -Property ResourcePool -Server $vc).ResourcePool
        }
    }

    # 目的資料夾
    $targetFolder = Resolve-VmFolder $dc $Folder -Create:$CreateFolders
    if (-not $targetFolder) {
        if ($DryRun) { $targetFolder = Get-Folder -Id $dc.ExtensionData.VmFolder.ToString() -Server $vc }
        else { throw "資料夾 '$Folder' 不存在（加 -CreateFolders 可自動建）" }
    }
    $folderView = Get-View $targetFolder.ExtensionData.MoRef -Server $vc

    # 掃 datastore
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
    if ($task.Info.State -ne 'success') { throw "掃描 $dsName 失敗：$($task.Info.Error.LocalizedMessage)" }

    $found = New-Object System.Collections.ArrayList
    foreach ($r in $task.Info.Result) {
        foreach ($f in $r.File) {
            $ext = [IO.Path]::GetExtension($f.Path).ToLower()
            if ($ext -notin '.vmx', '.vmtx') { continue }
            # FolderPath 形式：根目錄 "[ds]"、子目錄 "[ds] dir/"；組成 "[ds] dir/file.vmx"（與 VmPathName 同格式）
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
    Write-Host "  找到 $($found.Count) 個 vmx/vmtx"

    $i = 0
    foreach ($item in ($found | Sort-Object Path)) {
        if ($registered.ContainsKey($item.Path.ToLower())) { Add-Result 'AlreadyRegistered' $item.Path; continue }
        if (-not (Test-NameMatch $item.Name))              { Add-Result 'Filtered' $item.Path; continue }
        if ($item.IsTemplate -and $NoTemplates)            { Add-Result 'SkippedTemplate' $item.Path; continue }

        $h = $hosts[$i % $hosts.Count]; $i++
        $kind = if ($item.IsTemplate) { 'Template' } else { 'VM' }

        # VM 名稱：預設抓 vmx 內 displayName（RegisterVM 的 name 不能空），讀不到才用檔名
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

            # 依 placement 搬到各自資料夾
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

# --- 照舊 vC 補自訂屬性 / Notes / tag（資料夾在註冊時已依 placement 放好）---
if ($MetaDir -and $script:RegisteredNames.Count) {
    Write-Host "`n==================== 依舊 vC 補自訂屬性 / Notes / tag（$($script:RegisteredNames.Count) 台）===================="
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

# --- 對帳：匯出清單裡、在這些 datastore 上、但 B 端還沒出現的 VM ---
if ($PlacementCsv -and (Test-Path $PlacementCsv)) {
    $manifest = @(Import-Csv $PlacementCsv)
    if ($manifest.Count -and ($manifest[0].PSObject.Properties.Name -contains 'VmPathName')) {
        $onTarget = @{}
        foreach ($v in (Get-View -ViewType VirtualMachine -Property Name -Server $vc)) { $onTarget[$v.Name] = $true }
        if (-not $DryRun) { foreach ($n in $script:RegisteredNames) { $onTarget[$n] = $true } }
        # 期望集合：優先用 Unregister-VmFromOldVc 寫的 unregistered.csv（實際拔掉的那批），
        # 沒有就用匯出清單裡 vmx 在這些 datastore 上的
        $unregLog = Join-Path (Split-Path $PlacementCsv -Parent) 'unregistered.csv'
        if (Test-Path $unregLog) {
            $expected = @(Import-Csv $unregLog | Where-Object { $p = $_.VmPathName; $p -and ($Datastore | Where-Object { $p.StartsWith("[$_] ") }) })
            $basis = 'unregistered.csv'
        } else {
            $expected = @($manifest | Where-Object { $p = $_.VmPathName; $p -and ($Datastore | Where-Object { $p.StartsWith("[$_] ") }) })
            $basis = 'vm-placement.csv'
        }
        $pending = @($expected | Where-Object { -not $onTarget.ContainsKey($_.VMName) })
        Write-Host "`n================ 對帳（$basis vs 目標端）================"
        if ($DryRun) {
            $covered = @($pending | Where-Object { $script:RegisteredNames -contains $_.VMName }).Count
            Write-Host ("  舊 vC 端位於 {0} 的 VM：{1} 台；目標端已有 {2} 台；還沒過來 {3} 台（這次會註冊 {4} 台，跑完剩 {5} 台）" -f ($Datastore -join ","), $expected.Count, ($expected.Count - $pending.Count), $pending.Count, $covered, ($pending.Count - $covered))
        } else {
            Write-Host ("  舊 vC 端位於 {0} 的 VM：{1} 台；目標端已有 {2} 台；還沒過來 {3} 台" -f ($Datastore -join ","), $expected.Count, ($expected.Count - $pending.Count), $pending.Count)
        }
        foreach ($m in $pending) {
            Add-Result 'PendingOnSource' $m.VMName "$($m.VmPathName)"
        }
    }
}

# --- 總結 ---
Write-Host "`n================ 結果 ================"
foreach ($k in ($script:Stats.Keys | Sort-Object)) { Write-Host ("  {0,-20} {1,5}" -f $k, $script:Stats[$k]) }
if ($script:Report.Count) {
    $script:Report | Export-Csv -Path $ReportPath -NoTypeInformation -Encoding $enc
    Write-Host "`n[+] 明細報告 -> $ReportPath"
}
Disconnect-VIServer -Server $vc -Confirm:$false | Out-Null
