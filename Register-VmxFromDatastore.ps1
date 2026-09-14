<#
.SYNOPSIS
  掃指定 datastore 裡的 .vmx / .vmtx，把還沒在 inventory 裡的 VM / 範本註冊回 vCenter。

  典型用途：datastore 從 vCenter A 搬到 vCenter B 後，VM 檔案都在、但 inventory 是空的。
  註冊完再用 Import-VcMeta.ps1 補資料夾 / tag / 自訂屬性 / Notes。

.EXAMPLE
  # 先看會註冊哪些（不寫入）
  ./Register-VmxFromDatastore.ps1 -Server <vC> -User administrator@vsphere.local -Password '<pw>' -Datastore ds01 -Cluster cl01 -DryRun

  # 正式註冊，並依 vm-placement.csv 直接放進對應資料夾
  ./Register-VmxFromDatastore.ps1 -Server <vC> -User administrator@vsphere.local -Password '<pw>' -Datastore ds01 -Cluster cl01 -PlacementCsv .\export-A\vm-placement.csv

.NOTES
  - 已註冊的 vmx（依 [datastore] 路徑比對）一律跳過，重複跑安全。
  - .vmtx 會註冊成範本（-NoTemplates 可略過）。
  - 只註冊、不開機。開機時 vSphere 若問「moved / copied」要另外回答。
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$Server,
    [string]$User,
    [string]$Password,
    [pscredential]$Credential,

    [Parameter(Mandatory)][string[]]$Datastore,
    [string]$Cluster,                  # 註冊到這個叢集（挑有掛該 datastore 的主機，輪流用）
    [string]$VMHost,                   # 或指定單一主機
    [string]$ResourcePool,             # 預設用主機所屬叢集 / 主機的根 resource pool

    [string]$Folder,                   # 註冊到哪個 VM 資料夾（相對 datacenter 的路徑，例 'Linux'），預設 DC 根
    [string]$PlacementCsv,             # Export-VcMeta 的 vm-placement.csv：註冊後依 VMName 搬到各自資料夾
    [switch]$CreateFolders,            # 資料夾不存在就建

    [string[]]$Include,                # 只註冊符合的名稱（wildcard，例 'web*','db01'）
    [string[]]$Exclude = @('vCLS*'),   # 排除（預設略過 vCLS）
    [switch]$NameFromFile,             # 用 vmx 檔名當 VM 名稱（預設抓 vmx 裡的 displayName，讀不到才用檔名）
    [switch]$NoTemplates,
    [switch]$DryRun,
    [string]$ReportPath
)

$ErrorActionPreference = 'Stop'
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
foreach ($dsName in $Datastore) {
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
        if ($DryRun) { Add-Result 'WouldRegister' $item.Path "$kind -> host=$($h.Name) folder=$($targetFolder.Name)"; continue }

        try {
            $name = $null
            if (-not $NameFromFile) {
                $rel = ($item.Path -replace '^\[[^\]]+\]\s*', '')
                $name = Get-VmxDisplayName $dc.Name $ds.Name $rel
            }
            if ([string]::IsNullOrWhiteSpace($name)) { $name = $item.Name }   # 讀不到就用檔名
            $pool = if ($item.IsTemplate) { $null } else { $poolByHost[$h.Name] }
            $vmRef = $folderView.RegisterVM($item.Path, $name, $item.IsTemplate, $pool, $h.ExtensionData.MoRef)
            $registered[$item.Path.ToLower()] = $true
            $vmObj = Get-VIObjectByVIView -MORef $vmRef -Server $vc
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

# --- 總結 ---
Write-Host "`n================ 結果 ================"
foreach ($k in ($script:Stats.Keys | Sort-Object)) { Write-Host ("  {0,-20} {1,5}" -f $k, $script:Stats[$k]) }
if ($script:Report.Count) {
    $script:Report | Export-Csv -Path $ReportPath -NoTypeInformation -Encoding $enc
    Write-Host "`n[+] 明細報告 -> $ReportPath"
}
Disconnect-VIServer -Server $vc -Confirm:$false | Out-Null
