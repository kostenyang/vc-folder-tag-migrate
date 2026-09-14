<#
.SYNOPSIS
  舊 vC 端：依匯出清單（vm-placement.csv）把 VM / 範本從 inventory 移除（檔案留在 datastore）。
  這是「落地」流程的第 ② 步，之後 datastore 搬到新 vC 再用 Register-VmxFromDatastore -MetaDir 註冊回來。

  安全機制：
    - 沒有 -MetaDir（先跑過 Export-VcMeta）不給做 —— unregister 之後 tag 指派 / 屬性值就沒了
    - 清單裡的 vmx 路徑要跟現在一致（確保匯出檔不是舊的）
    - 開機中的 VM 不動（或 -ShutdownFirst）
    - 每台寫進 <MetaDir>\unregistered.csv（名稱 / vmx / 時間），新 vC 端對帳用

.EXAMPLE
  # 把 ds01 上、清單裡的 VM 全部 unregister（先看）
  ./Unregister-VmFromOldVc.ps1 -Server vcA -Password 'x' -MetaDir .\export-A -Datastore ds01 -DryRun

  # 只做某個資料夾的
  ./Unregister-VmFromOldVc.ps1 -Server vcA -Password 'x' -MetaDir .\export-A -Folder 'Linux'
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$Server,
    [string]$User = 'administrator@vsphere.local',
    [Parameter(Mandatory)][string]$Password,
    [Parameter(Mandatory)][string]$MetaDir,        # Export-VcMeta 的 -OutDir

    [string[]]$Datastore,              # 清單裡 vmx 在這些 datastore 上的
    [string[]]$Folder,                 # 清單裡在這些資料夾子樹的
    [string[]]$VM,                     # 或直接點名
    [switch]$ShutdownFirst,
    [int]$ShutdownTimeoutSec = 300,
    [switch]$DryRun,
    [string]$ReportPath
)

$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding $false } catch { }
$enc = if ($PSVersionTable.PSVersion.Major -ge 6) { 'utf8BOM' } else { 'UTF8' }
if (-not $Datastore -and -not $Folder -and -not $VM) { throw "要給 -Datastore / -Folder / -VM 至少一個，不接受「全部」" }
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
    $tag = switch -Wildcard ($Action) { 'Unregistered' { '[+]' } 'Would*' { '[~]' } 'Failed*' { '[X]' } 'Skipped*' { '[-]' } default { '[ ]' } }
    Write-Host ("  {0} {1,-18} {2}  {3}" -f $tag, $Action, $Target, $Detail)
}

# --- 從清單挑要做的 ---
$picked = @($manifest | Where-Object {
    $r = $_
    $ok = $false
    if ($VM        -and ($VM -contains $r.VMName)) { $ok = $true }
    if ($Datastore -and ($Datastore | Where-Object { $r.VmPathName.StartsWith("[$_] ") })) { $ok = $true }
    if ($Folder    -and ($Folder | Where-Object { $f = $_.Trim('/'); $r.FolderPath -eq $f -or $r.FolderPath -like "$f/*" })) { $ok = $true }
    $ok
})
Write-Host "[*] 清單 $($manifest.Count) 台，符合條件 $($picked.Count) 台"
if (-not $picked.Count) { return }

# --- 連線 ---
Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -DisplayDeprecationWarnings $false -Scope Session -Confirm:$false | Out-Null
$vc = Connect-VIServer -Server $Server -User $User -Password $Password
Write-Host "[+] 已連線 $($vc.Name) ($($vc.Version))"
if ($DryRun) { Write-Host "[!] DryRun：只檢查、不 unregister" -ForegroundColor Yellow }

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
        [void]$done.Add([pscustomobject]@{ Time = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'); VMName = $r.VMName; IsTemplate = $isTpl; VmPathName = $curVmx; FolderPath = $r.FolderPath; InstanceUuid = $r.InstanceUuid; Server = $Server })
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
