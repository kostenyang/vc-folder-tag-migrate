<#
.SYNOPSIS
  舊 vC 端：依匯出清單（vm-placement.csv）把 VM / 範本從 inventory 移除（檔案留在 datastore）。
  這是「落地」流程的第 ② 步，之後 datastore 搬到新 vC 再用 Register-VmxFromDatastore -MetaDir 註冊回來。

  安全機制：
    - 沒有 -MetaDir（先跑過 Export-VcMeta）不給做 —— unregister 之後 tag 指派 / 屬性值就沒了
    - 清單裡的 vmx 路徑要跟現在一致（確保匯出檔不是舊的）
    - 開機中的 VM 不動（或 -ShutdownFirst）
    - -Datastore / -Lun 時檢查跨 datastore 的 VM：名單內有 vmdk 在別顆（WarnDiskElsewhere）、
      名單外卻有 vmdk 在這顆（WarnDiskOnDatastore）—— 只警告不擋，請先搬碟或一起搬
    - 每台寫進 <MetaDir>\unregistered.csv（名稱 / vmx / 原資料夾 / 時間）——這份就是下一步\n      Register-VmxFromDatastore -UnregisteredCsv 的輸入，可先用 Excel 改（例如改目的資料夾）再註冊

.EXAMPLE
  # 把 ds01 上、清單裡的 VM 全部 unregister（先看）
  ./Unregister-VmFromOldVc.ps1 -Server vcA -Password 'x' -MetaDir .\export-A -Datastore ds01 -DryRun

  # cluster 裡且在 ds01 上的（多條件取交集）
  ./Unregister-VmFromOldVc.ps1 -Server vcA -Password 'x' -MetaDir .\export-A -Cluster cl01 -Datastore ds01

  # 儲存團隊給的是 LUN ID：反查落在這顆 LUN 上的 datastore（可只給尾碼）
  ./Unregister-VmFromOldVc.ps1 -Server vcA -Password 'x' -MetaDir .\export-A -Lun naa.6000c29b1a2b3c4d

  # 只做某個資料夾的
  ./Unregister-VmFromOldVc.ps1 -Server vcA -Password 'x' -MetaDir .\export-A -Folder 'Linux'
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$Server,
    [string]$User = 'administrator@vsphere.local',
    [Parameter(Mandatory)][string]$Password,
    [Parameter(Mandatory)][string]$MetaDir,        # Export-VcMeta 的 -OutDir

    # 篩選條件：可以給多個，同時給時取「交集」（例：-Cluster cl01 -Datastore ds01 = cl01 裡且在 ds01 上的）
    [string[]]$Datastore,              # 清單裡 vmx 在這些 datastore 上的
    [string[]]$Lun,                    # 底層 LUN（naa.xxx / eui.xxx / t10.xxx，可只給尾碼、可 wildcard）→ 反查落在上面的 VMFS datastore
    [string[]]$Cluster,                # 目前跑在這些叢集上的（連舊 vC 即時查）
    [string[]]$VMHost,                 # 目前跑在這些主機上的（連舊 vC 即時查）
    [string[]]$Folder,                 # 清單裡在這些資料夾子樹的
    [string[]]$VM,                     # 直接點名
    [switch]$ShutdownFirst,
    [int]$ShutdownTimeoutSec = 300,
    [switch]$DryRun,
    [string]$ReportPath
)

$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding $false } catch { }
$enc = if ($PSVersionTable.PSVersion.Major -ge 6) { 'utf8BOM' } else { 'UTF8' }
if (-not $Datastore -and -not $Lun -and -not $Cluster -and -not $VMHost -and -not $Folder -and -not $VM) { throw "要給 -Datastore / -Lun / -Cluster / -VMHost / -Folder / -VM 至少一個，不接受「全部」" }
$MetaDir = (Resolve-Path $MetaDir).Path
$manifestPath = Join-Path $MetaDir 'vm-placement.csv'
if (-not (Test-Path $manifestPath)) { throw "$MetaDir 裡沒有 vm-placement.csv —— 先跑 Export-VcMeta.ps1 再來" }
$manifest = @(Import-Csv $manifestPath)
if (-not $manifest.Count -or ($manifest[0].PSObject.Properties.Name -notcontains 'VmPathName')) { throw "vm-placement.csv 太舊（沒有 VmPathName 欄位），請用新版 Export-VcMeta 重新匯出" }
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
    Write-Host ("  {0} {1,-18} {2}  {3}" -f $tag, $Action, $Target, $Detail)
}

# --- 連線（-Cluster / -VMHost 要即時查舊 vC）---
Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -DisplayDeprecationWarnings $false -Scope Session -Confirm:$false | Out-Null
$vc = Connect-VIServer -Server $Server -User $User -Password $Password
Write-Host "[+] 已連線 $($vc.Name) ($($vc.Version))"
if ($DryRun) { Write-Host "[!] DryRun：只檢查、不 unregister" -ForegroundColor Yellow }

# --- -Lun：用 VMFS extent 反查 datastore，併進 -Datastore（同一個維度，取聯集）---
if ($Lun) {
    $hit = @()
    foreach ($dsv in (Get-View -ViewType Datastore -Property Name,Info -Server $vc)) {
        $ext = @(); try { if ($dsv.Info.Vmfs) { $ext = @($dsv.Info.Vmfs.Extent | ForEach-Object { $_.DiskName }) } } catch { }
        foreach ($pat in $Lun) {
            $pp = if ($pat -match '[*?]') { $pat } else { "*$pat*" }
            if ($ext | Where-Object { $_ -like $pp }) { $hit += $dsv.Name; Write-Host "  [*] LUN $pat → datastore $($dsv.Name)（$($ext -join ';')）"; break }
        }
    }
    if (-not $hit) { Write-Host "  [!] 舊 vC 上沒有 datastore 落在 LUN $($Lun -join ',')"; Disconnect-VIServer -Server $vc -Confirm:$false | Out-Null; return }
    $Datastore = @(@($Datastore) + $hit | Where-Object { $_ } | Sort-Object -Unique)
}

# --- 從清單挑要做的（多個條件取交集）---
$inCluster = $null
if ($Cluster -or $VMHost) {
    $inCluster = @{}
    $hostIds = @{}
    foreach ($n in $Cluster) {
        $o = Get-Cluster -Name $n -Server $vc -ErrorAction SilentlyContinue
        if (-not $o) { Write-Host "  [!] 舊 vC 找不到叢集 '$n'"; continue }
        foreach ($h in ($o | Get-VMHost -Server $vc)) { $hostIds[$h.ExtensionData.MoRef.ToString()] = $true }
    }
    foreach ($n in $VMHost) {
        $o = Get-VMHost -Name $n -Server $vc -ErrorAction SilentlyContinue
        if (-not $o) { Write-Host "  [!] 舊 vC 找不到主機 '$n'"; continue }
        $hostIds[$o.ExtensionData.MoRef.ToString()] = $true
    }
    # 用 view 依 Runtime.Host 篩：VM 與範本一次涵蓋（Get-Template -Location 不接受 Cluster）
    foreach ($v in (Get-View -ViewType VirtualMachine -Property Name,Runtime.Host -Server $vc)) {
        if ($v.Runtime.Host -and $hostIds.ContainsKey($v.Runtime.Host.ToString())) { $inCluster[$v.Name] = $true }
    }
    Write-Host "[*] 叢集/主機 $((@($Cluster) + @($VMHost) | Where-Object { $_ }) -join ',') 上目前有 $($inCluster.Count) 台"
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
if ($VM) { $cond += "VM=$($VM -join ',')" }; if ($Datastore) { $cond += "datastore=$($Datastore -join ',')" }; if ($Lun) { $cond += "lun=$($Lun -join ',')" }
if ($Cluster) { $cond += "cluster=$($Cluster -join ',')" }; if ($VMHost) { $cond += "host=$($VMHost -join ',')" }; if ($Folder) { $cond += "folder=$($Folder -join ',')" }
Write-Host "[*] 清單 $($manifest.Count) 台，符合 [$($cond -join ' AND ')] 的 $($picked.Count) 台"
if (-not $picked.Count) { Disconnect-VIServer -Server $vc -Confirm:$false | Out-Null; return }

# --- 跨 datastore 檢查（只有 -Datastore / -Lun 時有意義）---
# 只搬一顆 LUN 的話，這兩種 VM 會壞：
#   A) 在名單內（vmx 在目標 datastore）但有 vmdk 在別顆      → 註冊到新 vC 後那顆碟找不到
#   B) 不在名單內（vmx 在別顆）但有 vmdk 在目標 datastore    → 留在舊 vC 的 VM 碟被搬走
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
            Add-Result 'WarnDiskElsewhere' $v.Name "有 vmdk 在 $($diskOut -join ',')，只搬 $($Datastore -join ',') 這顆會缺碟"
        }
        if (-not $homeIn -and $diskIn.Count) {
            Add-Result 'WarnDiskOnDatastore' $v.Name "vmx 在 $homeDs、但有 vmdk 在 $($diskIn -join ',')，不在名單內卻會被搬走"
        }
    }
}

$done = New-Object System.Collections.ArrayList
foreach ($r in $picked) {
    $o = Get-VM -Name $r.VMName -Server $vc -ErrorAction SilentlyContinue
    $isTpl = $false
    if (-not $o) { $o = Get-Template -Name $r.VMName -Server $vc -ErrorAction SilentlyContinue; $isTpl = [bool]$o }
    if (-not $o) { Add-Result 'SkippedNotFound' $r.VMName '舊 vC 上已經沒有（可能已 unregister）'; continue }
    if (@($o).Count -gt 1) { Add-Result 'SkippedAmbiguous' $r.VMName '同名多台，請用資料夾或 UUID 區分'; continue }

    $curVmx = $o.ExtensionData.Config.Files.VmPathName
    if ($curVmx -ne $r.VmPathName) { Add-Result 'SkippedStale' $r.VMName "清單 $($r.VmPathName) ≠ 現在 $curVmx，請重新匯出"; continue }

    if (-not $isTpl -and $o.PowerState -ne 'PoweredOff') {
        if (-not $ShutdownFirst) { Add-Result 'SkippedPoweredOn' $r.VMName "$($o.PowerState)，先關機或加 -ShutdownFirst"; continue }
        if ($DryRun) { Add-Result 'WouldShutdown' $r.VMName }
        else {
            try { $null = Shutdown-VMGuest -VM $o -Confirm:$false -Server $vc } catch { Add-Result 'FailedShutdown' $r.VMName ($_.Exception.Message -split "`n")[0]; continue }
            $deadline = (Get-Date).AddSeconds($ShutdownTimeoutSec)
            do { Start-Sleep 5; $o = Get-VM -Name $r.VMName -Server $vc } while ($o.PowerState -ne 'PoweredOff' -and (Get-Date) -lt $deadline)
            if ($o.PowerState -ne 'PoweredOff') { Add-Result 'FailedShutdown' $r.VMName "等 $ShutdownTimeoutSec 秒還沒關"; continue }
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
    Write-Host "`n[+] 已 unregister 的 $($done.Count) 台記在 $logPath（新 vC 端對帳用）"
}
Write-Host "`n================ 結果 ================"
foreach ($k in ($script:Stats.Keys | Sort-Object)) { Write-Host ("  {0,-22} {1,5}" -f $k, $script:Stats[$k]) }
$script:Report | Export-Csv -Path $ReportPath -NoTypeInformation -Encoding $enc
Write-Host "[+] 明細報告 -> $ReportPath"
Disconnect-VIServer -Server $vc -Confirm:$false | Out-Null
