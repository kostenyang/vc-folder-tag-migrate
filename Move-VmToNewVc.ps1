<#
.SYNOPSIS
  VM 從舊 vCenter 交接到新 vCenter，中間不會有「兩邊都看不到」的空窗。
  前提：VM 所在的 datastore 兩台 vC 都掛得到（共用 NFS / SAN LUN 兩邊都 present）。

  每台 VM 依序：
    ① 確認關機（或 -ShutdownFirst 幫你關）
    ② 新 vC 先註冊（同一個 vmx，關機狀態下兩邊短暫都看得到是安全的）
    ③ 註冊成功 → 才從舊 vC unregister（失敗就不動舊 vC，VM 還在原地）
    ④ 放回原資料夾、補自訂屬性 / Notes / tag（Export 在一開始自動做，那時 VM 還在舊 vC）
    ⑤ -PowerOnAfter：新 vC 開機，自動回答「I moved it」

.EXAMPLE
  # 先看會動哪些
  ./Move-VmToNewVc.ps1 -SourceServer vcA -SourcePassword 'x' -TargetServer vcB -TargetPassword 'y' -Folder 'Linux' -TargetCluster cl01 -DatacenterMap 'DC-A=DC-B' -DryRun

  # 指定幾台，關機→交接→開機一條龍
  ./Move-VmToNewVc.ps1 ... -VM web01,web02 -TargetCluster cl01 -ShutdownFirst -PowerOnAfter
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$SourceServer,
    [string]$SourceUser = 'administrator@vsphere.local',
    [Parameter(Mandatory)][string]$SourcePassword,
    [Parameter(Mandatory)][string]$TargetServer,
    [string]$TargetUser = 'administrator@vsphere.local',
    [Parameter(Mandatory)][string]$TargetPassword,

    [string[]]$VM,                     # 要搬的 VM 名稱（可多個）
    [string[]]$Folder,                 # 或整個資料夾子樹（舊 vC 的相對路徑）
    [string]$TargetCluster,            # 新 vC 落點：叢集（挑有掛該 datastore 的主機輪流用）
    [string]$TargetVMHost,             # 或指定主機
    [string[]]$DatacenterMap,          # 'DC-A=DC-B'

    [string]$MetaDir,                  # 匯出目錄；不存在或沒有清單就自動在開始時 Export（VM 還在舊 vC）
    [switch]$ShutdownFirst,            # 開著的 VM 先 guest shutdown（要 VMware Tools）
    [int]$ShutdownTimeoutSec = 300,
    [switch]$PowerOnAfter,             # 交接完在新 vC 開機，並自動回答 moved
    [switch]$DryRun,
    [string]$ReportPath
)

$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding $false } catch { }
$enc = if ($PSVersionTable.PSVersion.Major -ge 6) { 'utf8BOM' } else { 'UTF8' }
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $VM -and -not $Folder) { throw "要給 -VM 或 -Folder 其中一個" }
if (-not $ReportPath) { $ReportPath = Join-Path (Get-Location) ('move-report-{0:yyyyMMdd-HHmmss}.csv' -f (Get-Date)) }

$script:Report = New-Object System.Collections.ArrayList
$script:Stats  = @{}
function Add-Result {
    param([string]$Action,[string]$Target,[string]$Detail = '')
    [void]$script:Report.Add([pscustomobject]@{ Time = (Get-Date -Format 'HH:mm:ss'); Action = $Action; Target = $Target; Detail = $Detail })
    if (-not $script:Stats.ContainsKey($Action)) { $script:Stats[$Action] = 0 }
    $script:Stats[$Action]++
    $tag = switch -Wildcard ($Action) { 'Moved' { '[+]' } 'Would*' { '[~]' } 'Failed*' { '[X]' } 'Skipped*' { '[-]' } default { '[ ]' } }
    Write-Host ("  {0} {1,-18} {2}  {3}" -f $tag, $Action, $Target, $Detail)
}
function Map-Dc { param([string]$Name); foreach ($m in $DatacenterMap) { $kv = $m -split '=', 2; if ($kv.Count -eq 2 -and $kv[0] -eq $Name) { return $kv[1] } }; return $Name }

# ============ 0. 匯出（VM 還在舊 vC 的時候）============
if (-not $MetaDir) { $MetaDir = Join-Path (Get-Location) ('move-meta-{0}-{1:yyyyMMdd-HHmmss}' -f ($SourceServer -replace '[^\w.-]','_'), (Get-Date)) }
if (-not (Test-Path (Join-Path $MetaDir 'vm-placement.csv'))) {     # 沒有現成匯出檔就現在匯（VM 還在舊 vC）
    Write-Host "==================== [0] 從舊 vC $SourceServer 匯出資料夾 / tag / 屬性 / Notes ===================="
    $exp = @{ Server = $SourceServer; User = $SourceUser; Password = $SourcePassword; OutDir = $MetaDir }
    if ($Folder -and -not $VM) { $exp.Folder = $Folder }
    & "$here\Export-VcMeta.ps1" @exp
    Write-Host ""
}
$MetaDir = (Resolve-Path $MetaDir).Path
$placement = @{}
foreach ($r in (Import-Csv (Join-Path $MetaDir 'vm-placement.csv'))) { $placement[$r.VMName] = $r }

# ============ 1. 同時連兩台 ============
Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -DisplayDeprecationWarnings $false -DefaultVIServerMode Multiple -Scope Session -Confirm:$false | Out-Null
Write-Host "[*] 連線舊 vC $SourceServer / 新 vC $TargetServer ..."
$src = Connect-VIServer -Server $SourceServer -User $SourceUser -Password $SourcePassword
$tgt = Connect-VIServer -Server $TargetServer -User $TargetUser -Password $TargetPassword
Write-Host "[+] 舊 $($src.Name) ($($src.Version))  →  新 $($tgt.Name) ($($tgt.Version))"
if ($DryRun) { Write-Host "[!] DryRun 模式：只檢查、不註冊、不 unregister" -ForegroundColor Yellow }

# ============ 2. 挑 VM（舊 vC）============
$targets = New-Object System.Collections.ArrayList
if ($VM) {
    foreach ($n in $VM) {
        $o = Get-VM -Name $n -Server $src -ErrorAction SilentlyContinue
        if (-not $o) { $o = Get-Template -Name $n -Server $src -ErrorAction SilentlyContinue }
        if ($o) { [void]$targets.Add($o) } else { Add-Result 'NotFoundOnSource' $n }
    }
} else {
    foreach ($fp in $Folder) {
        $dcs = Get-Datacenter -Server $src
        $found = $false
        foreach ($dc in $dcs) {
            $cur = Get-Folder -Id $dc.ExtensionData.VmFolder.ToString() -Server $src
            foreach ($seg in ($fp.Trim('/') -split '/')) { $cur = Get-Folder -Location $cur -Name $seg -NoRecursion -Server $src -ErrorAction SilentlyContinue | Select-Object -First 1; if (-not $cur) { break } }
            if ($cur) {
                $found = $true
                foreach ($o in (Get-VM -Location $cur -Server $src)) { [void]$targets.Add($o) }
                foreach ($o in (Get-Template -Location $cur -Server $src)) { [void]$targets.Add($o) }
            }
        }
        if (-not $found) { Add-Result 'FolderNotFoundOnSource' $fp }
    }
}
Write-Host "[*] 要交接 $($targets.Count) 台"

# ============ 3. 新 vC 端的工具 ============
$script:FolderCache = @{}
function Resolve-TargetFolder {
    param([string]$DcName, [string]$Path)
    $dc = Get-Datacenter -Name $DcName -Server $tgt -ErrorAction SilentlyContinue
    if (-not $dc) { return $null }
    $parent = Get-Folder -Id $dc.ExtensionData.VmFolder.ToString() -Server $tgt
    if ([string]::IsNullOrEmpty($Path)) { return $parent }
    $acc = @()
    foreach ($seg in ($Path -split '/')) {
        $acc += $seg
        $k = "$DcName|$($acc -join '/')"
        if ($script:FolderCache.ContainsKey($k)) { $parent = $script:FolderCache[$k]; continue }
        $child = Get-Folder -Location $parent -Name $seg -NoRecursion -Server $tgt -ErrorAction SilentlyContinue | Select-Object -First 1
        if (-not $child) {
            if ($DryRun) { Add-Result 'WouldCreateFolder' $k; return $null }
            $child = New-Folder -Name $seg -Location $parent -Server $tgt
            Add-Result 'CreatedFolder' $k
        }
        $script:FolderCache[$k] = $child; $parent = $child
    }
    return $parent
}
$hostCache = @{}   # datastore 名 -> 可用主機清單
$rr = @{}
function Get-TargetHost {
    param([string]$DsName)
    if (-not $hostCache.ContainsKey($DsName)) {
        $ds = Get-Datastore -Name $DsName -Server $tgt -ErrorAction SilentlyContinue
        $hs = @()
        if ($ds) {
            $hs = @($ds | Get-VMHost -Server $tgt | Where-Object { $_.ConnectionState -eq 'Connected' })
            if ($TargetCluster) { $hs = @($hs | Where-Object { $_.Parent.Name -eq $TargetCluster }) }
            if ($TargetVMHost)  { $hs = @($hs | Where-Object { $_.Name -eq $TargetVMHost }) }
        }
        $hostCache[$DsName] = $hs; $rr[$DsName] = 0
    }
    $hs = $hostCache[$DsName]
    if (-not $hs.Count) { return $null }
    $h = $hs[$rr[$DsName] % $hs.Count]; $rr[$DsName]++
    return $h
}

# ============ 4. 逐台交接 ============
$moved = New-Object System.Collections.ArrayList
Write-Host "`n==================== 交接 ===================="
foreach ($o in $targets) {
    $name  = $o.Name
    $isTpl = ($o.GetType().Name -match 'Template')
    $vmx   = $o.ExtensionData.Config.Files.VmPathName
    $dsName = ($vmx -replace '^\[([^\]]+)\].*$', '$1')
    $rec = $placement[$name]
    $dstDc = if ($rec) { Map-Dc $rec.Datacenter } else { Map-Dc (Get-Datacenter -VM $o -Server $src).Name }
    $dstPath = if ($rec) { $rec.FolderPath } else { '' }

    # 新 vC 看得到這顆 datastore 嗎
    $h = Get-TargetHost $dsName
    if (-not $h) { Add-Result 'SkippedNoDatastore' $name "新 vC 沒有主機掛著 $dsName（Cluster=$TargetCluster Host=$TargetVMHost）"; continue }

    # 電源
    if (-not $isTpl -and $o.PowerState -ne 'PoweredOff') {
        if (-not $ShutdownFirst) { Add-Result 'SkippedPoweredOn' $name "$($o.PowerState)，加 -ShutdownFirst 或先手動關機"; continue }
        if ($DryRun) { Add-Result 'WouldShutdown' $name }
        else {
            try { $null = Shutdown-VMGuest -VM $o -Confirm:$false -Server $src } catch { Add-Result 'FailedShutdown' $name ($_.Exception.Message -split "`n")[0]; continue }
            $deadline = (Get-Date).AddSeconds($ShutdownTimeoutSec)
            do { Start-Sleep 5; $o = Get-VM -Name $name -Server $src } while ($o.PowerState -ne 'PoweredOff' -and (Get-Date) -lt $deadline)
            if ($o.PowerState -ne 'PoweredOff') { Add-Result 'FailedShutdown' $name "等 $ShutdownTimeoutSec 秒還沒關" ; continue }
            Add-Result 'Shutdown' $name
        }
    }

    $dstFolder = Resolve-TargetFolder $dstDc $dstPath
    if ($DryRun) {
        Add-Result 'WouldMove' $name "$vmx -> $($h.Name) / ${dstDc}:${dstPath}"
        [void]$moved.Add($name)
        continue
    }
    if (-not $dstFolder) { Add-Result 'FailedFolder' $name "新 vC 找不到 Datacenter '$dstDc'（用 -DatacenterMap）"; continue }

    # ② 新 vC 先註冊
    try {
        $fv = Get-View $dstFolder.ExtensionData.MoRef -Server $tgt
        $pool = if ($isTpl) { $null } else { (Get-View $h.ExtensionData.Parent -Property ResourcePool -Server $tgt).ResourcePool }
        $newRef = $fv.RegisterVM($vmx, $name, $isTpl, $pool, $h.ExtensionData.MoRef)
    } catch {
        Add-Result 'FailedRegister' $name ($_.Exception.Message -split "`n")[0]; continue     # 舊 vC 不動
    }
    # ③ 成功才從舊 vC 拿掉
    try {
        if ($isTpl) { Remove-Template -Template $o -Confirm:$false -Server $src } else { Remove-VM -VM $o -Confirm:$false -Server $src }
    } catch {
        Add-Result 'FailedUnregister' $name "新 vC 已註冊，但舊 vC unregister 失敗：$(($_.Exception.Message -split "`n")[0])"; [void]$moved.Add($name); continue
    }
    Add-Result 'Moved' $name "$vmx -> $($h.Name) / ${dstDc}:${dstPath}"
    [void]$moved.Add($name)

    # ⑤ 開機 + 回答 moved
    if ($PowerOnAfter -and -not $isTpl) {
        try {
            $nv = Get-VIObjectByVIView -MORef $newRef -Server $tgt
            $null = Start-VM -VM $nv -Confirm:$false -Server $tgt -RunAsync
            $deadline = (Get-Date).AddSeconds(90); $answered = $false
            do {
                Start-Sleep 3
                $q = Get-VMQuestion -VM $nv -Server $tgt -ErrorAction SilentlyContinue
                if ($q) { $q | Set-VMQuestion -Option 'button.uuid.movedTheVM' -Confirm:$false | Out-Null; $answered = $true }
                $nv = Get-VM -Id $nv.Id -Server $tgt
            } while ($nv.PowerState -ne 'PoweredOn' -and (Get-Date) -lt $deadline)
            Add-Result 'PoweredOn' $name $(if ($answered) { '已回答 I moved it' } else { '沒有詢問' })
        } catch { Add-Result 'FailedPowerOn' $name ($_.Exception.Message -split "`n")[0] }
    }
}

# ============ 5. 補自訂屬性 / Notes / tag ============
if ($moved.Count) {
    Write-Host "`n==================== 依舊 vC 補自訂屬性 / Notes / tag（$($moved.Count) 台）===================="
    $imp = @{ Server = $TargetServer; User = $TargetUser; Password = $TargetPassword; InDir = $MetaDir
              Include = @('CustomAttributes','Notes','Tags'); OnlyVMs = @($moved)
              ReportPath = ($ReportPath -replace '\.csv$', '-meta.csv') }
    if ($DatacenterMap) { $imp.DatacenterMap = $DatacenterMap }
    if ($DryRun)        { $imp.DryRun = $true }
    & "$here\Import-VcMeta.ps1" @imp
}

# ============ 總結 ============
Write-Host "`n================ 結果 ================"
foreach ($k in ($script:Stats.Keys | Sort-Object)) { Write-Host ("  {0,-22} {1,5}" -f $k, $script:Stats[$k]) }
if ($script:Report.Count) { $script:Report | Export-Csv -Path $ReportPath -NoTypeInformation -Encoding $enc; Write-Host "`n[+] 明細報告 -> $ReportPath" }
Write-Host "[+] 匯出檔在 $MetaDir（之後補跑 Register/Import 都用它）"
Disconnect-VIServer -Server $src, $tgt -Confirm:$false | Out-Null
