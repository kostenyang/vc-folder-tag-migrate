<#
.SYNOPSIS
  Register-VmxFromDatastore.ps1 的端到端自我測試：
  建一台丟棄式 VM + 一個範本 → 從 inventory 移除（檔案留著）→ 用腳本註冊回來
  → 驗證（含 placement 放資料夾、範本仍是範本、重跑冪等）→ 全部清掉。

.EXAMPLE
  ./Test-RegisterVmx.ps1 -Server 10.0.1.19 -Password '<pw>' -Datastore vcd-ds01 -Cluster m01-cl02
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$Server,
    [string]$User = 'administrator@vsphere.local',
    [Parameter(Mandatory)][string]$Password,
    [Parameter(Mandatory)][string]$Datastore,
    [string]$Cluster,
    [string]$Prefix = 'zz-regtest',
    [string]$WorkDir = "$env:TEMP\vc-regtest"
)

$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding $false } catch { }
Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -DisplayDeprecationWarnings $false -Scope Session -Confirm:$false | Out-Null

$here    = Split-Path -Parent $MyInvocation.MyCommand.Path
$vmName  = "$Prefix-vm01"
$tplName = "$Prefix-tpl01"
$fldTop  = "$Prefix-folder"
$fldPath = "$fldTop/Sub"
New-Item -ItemType Directory -Force -Path $WorkDir | Out-Null

$script:Checks = New-Object System.Collections.ArrayList
function Check {
    param([string]$Label, [bool]$Ok, [string]$Got)
    [void]$script:Checks.Add([pscustomobject]@{ Label = $Label; Ok = $Ok; Got = $Got })
    Write-Host ("  {0} {1,-40} {2}" -f $(if ($Ok) { '[PASS]' } else { '[FAIL]' }), $Label, $Got)
}
function Run-Register {
    # 一定用 hashtable splatting：陣列 splat 只做位置綁定，mandatory 參數沒綁到會卡在互動 prompt
    param([hashtable]$Extra = @{}, [string]$Report)
    $p = @{ Server = $Server; User = $User; Password = $Password; Datastore = $Datastore; ReportPath = $Report }
    if ($Cluster) { $p.Cluster = $Cluster }
    foreach ($k in $Extra.Keys) { $p[$k] = $Extra[$k] }
    & "$here\Register-VmxFromDatastore.ps1" @p | Out-Null
    @(Import-Csv $Report)
}

# ============ 1. 建測試 VM + 範本 ============
Write-Host "`n=== [1/6] 在 $Server 的 $Datastore 建測試 VM 與範本 ==="
$vc = Connect-VIServer -Server $Server -User $User -Password $Password
$ds = Get-Datastore -Name $Datastore -Server $vc
$dc = $ds.Datacenter
$host0 = $ds | Get-VMHost -Server $vc | Where-Object { $_.ConnectionState -eq 'Connected' -and (-not $Cluster -or $_.Parent.Name -eq $Cluster) } | Select-Object -First 1
if (-not $host0) { throw "沒有主機掛著 $Datastore" }

foreach ($n in $vmName, $tplName) {
    if (Get-VM -Name $n -Server $vc -ErrorAction SilentlyContinue)       { Remove-VM -VM $n -DeletePermanently -Confirm:$false -Server $vc }
    if (Get-Template -Name $n -Server $vc -ErrorAction SilentlyContinue) { Remove-Template -Template $n -DeletePermanently -Confirm:$false -Server $vc }
}
$vm  = New-VM -Name $vmName  -VMHost $host0 -Datastore $ds -NumCpu 1 -MemoryMB 128 -DiskMB 16 -GuestId otherLinux64Guest -Server $vc
$tv  = New-VM -Name $tplName -VMHost $host0 -Datastore $ds -NumCpu 1 -MemoryMB 128 -DiskMB 16 -GuestId otherLinux64Guest -Server $vc
$vmxPath  = $vm.ExtensionData.Config.Files.VmPathName
$tpl = Set-VM -VM $tv -ToTemplate -Confirm:$false -Server $vc
$tplPath  = (Get-Template -Name $tplName -Server $vc).ExtensionData.Config.Files.VmPathName
Write-Host "  VM     : $vmxPath"
Write-Host "  Template: $tplPath"

# ============ 2. 從 inventory 移除（檔案留著）============
Write-Host "`n=== [2/6] unregister（不刪檔）==="
Remove-VM -VM $vm -Confirm:$false -Server $vc                          # 沒有 -DeletePermanently = 只移出 inventory
Remove-Template -Template $tplName -Confirm:$false -Server $vc
$gone = -not (Get-VM -Name $vmName -Server $vc -ErrorAction SilentlyContinue) -and -not (Get-Template -Name $tplName -Server $vc -ErrorAction SilentlyContinue)
Check 'unregister 後 inventory 沒有它們' $gone ''
Disconnect-VIServer -Server $vc -Confirm:$false | Out-Null

# placement csv：只給 VM 指定資料夾，範本留在 DC 根
$csv = Join-Path $WorkDir 'vm-placement.csv'
@(
    [pscustomobject]@{ Datacenter = $dc.Name; VMName = $vmName; InstanceUuid = ''; BiosUuid = ''; IsTemplate = 'False'; FolderPath = $fldPath; InVApp = 'False'; MoRef = '' }
) | Export-Csv -Path $csv -NoTypeInformation -Encoding utf8

# ============ 3. DryRun ============
Write-Host "`n=== [3/6] DryRun ==="
$dry = Run-Register -Extra @{ DryRun = $true } -Report (Join-Path $WorkDir 'dry.csv')
$would = @($dry | Where-Object { $_.Action -eq 'WouldRegister' })
Check 'DryRun: 只會註冊這 2 個' (($would.Count -eq 2) -and (@($would | Where-Object { $_.Target -notlike "*$Prefix*" }).Count -eq 0)) "WouldRegister=$($would.Count)"
Check 'DryRun: 分得出 VM / Template' ((@($would | Where-Object { $_.Detail -like 'VM *' }).Count -eq 1) -and (@($would | Where-Object { $_.Detail -like 'Template *' }).Count -eq 1)) (($would.Detail | ForEach-Object { ($_ -split ' ')[0] }) -join ', ')
Check 'DryRun: 沒有真的註冊' (@($dry | Where-Object { $_.Action -eq 'Registered' }).Count -eq 0) ''

# ============ 4. 正式註冊 + placement ============
Write-Host "`n=== [4/6] 正式註冊（-PlacementCsv -CreateFolders）==="
$apply = Run-Register -Extra @{ PlacementCsv = $csv; CreateFolders = $true } -Report (Join-Path $WorkDir 'apply.csv')
Check '註冊: Registered = 2' (@($apply | Where-Object { $_.Action -eq 'Registered' }).Count -eq 2) ("Registered=" + @($apply | Where-Object { $_.Action -eq 'Registered' }).Count)
Check '註冊: Failed = 0'     (@($apply | Where-Object { $_.Action -eq 'Failed' }).Count -eq 0) (($apply | Where-Object { $_.Action -eq 'Failed' } | ForEach-Object { $_.Detail }) -join '; ')
Check '註冊: 依 placement 放資料夾' (@($apply | Where-Object { $_.Action -eq 'Placed' -and $_.Target -eq $vmName }).Count -eq 1) ''

# ============ 5. 獨立驗證 ============
Write-Host "`n=== [5/6] 目標端獨立驗證 ==="
$vc = Connect-VIServer -Server $Server -User $User -Password $Password
$v2 = Get-VM -Name $vmName -Server $vc -ErrorAction SilentlyContinue
$t2 = Get-Template -Name $tplName -Server $vc -ErrorAction SilentlyContinue
Check 'VM 回到 inventory、vmx 路徑一致' ($v2 -and $v2.ExtensionData.Config.Files.VmPathName -eq $vmxPath) $(if ($v2) { $v2.ExtensionData.Config.Files.VmPathName } else { '找不到' })
Check '範本回到 inventory 且仍是範本'   ($t2 -and $t2.ExtensionData.Config.Template -and $t2.ExtensionData.Config.Files.VmPathName -eq $tplPath) $(if ($t2) { "template=$($t2.ExtensionData.Config.Template)" } else { '找不到' })
$vmFolder = if ($v2) { $v2.Folder } else { $null }
$folderOk = $vmFolder -and $vmFolder.Name -eq 'Sub' -and $vmFolder.Parent.Name -eq $fldTop
Check 'VM 在 placement 指定的資料夾' $folderOk $(if ($vmFolder) { "$($vmFolder.Parent.Name)/$($vmFolder.Name)" } else { '' })
Check 'VM 沒被開機' ($v2 -and $v2.PowerState -eq 'PoweredOff') "$($v2.PowerState)"
Disconnect-VIServer -Server $vc -Confirm:$false | Out-Null

Write-Host "`n=== [5b] 重跑一次（冪等）==="
$again = Run-Register -Extra @{ PlacementCsv = $csv } -Report (Join-Path $WorkDir 'again.csv')
Check '重跑: 全部 AlreadyRegistered' ((@($again | Where-Object { $_.Action -eq 'Registered' }).Count -eq 0) -and (@($again | Where-Object { $_.Action -eq 'AlreadyRegistered' -and $_.Target -like "*$Prefix*" }).Count -eq 2)) ("Registered=" + @($again | Where-Object { $_.Action -eq 'Registered' }).Count)

# ============ 6. 清除 ============
Write-Host "`n=== [6/6] 清除 ==="
$vc = Connect-VIServer -Server $Server -User $User -Password $Password
if (Get-VM -Name $vmName -Server $vc -ErrorAction SilentlyContinue)        { Remove-VM -VM $vmName -DeletePermanently -Confirm:$false -Server $vc }
if (Get-Template -Name $tplName -Server $vc -ErrorAction SilentlyContinue) { Remove-Template -Template $tplName -DeletePermanently -Confirm:$false -Server $vc }
$root = Get-Folder -Id $dc.ExtensionData.VmFolder.ToString() -Server $vc
$top  = Get-Folder -Location $root -Name $fldTop -NoRecursion -Server $vc -ErrorAction SilentlyContinue
if ($top) { Remove-Folder -Folder $top -DeletePermanently -Confirm:$false -Server $vc }
$left = @()
if (Get-VM -Name $vmName -Server $vc -ErrorAction SilentlyContinue)        { $left += 'vm' }
if (Get-Template -Name $tplName -Server $vc -ErrorAction SilentlyContinue) { $left += 'template' }
if (Get-Folder -Name $fldTop -Server $vc -ErrorAction SilentlyContinue)    { $left += 'folder' }
# datastore 上的目錄也要不見
$browser = Get-View $ds.ExtensionData.Browser -Server $vc
$spec = New-Object VMware.Vim.HostDatastoreBrowserSearchSpec; $spec.MatchPattern = @("$Prefix*")
$t = Get-View ($browser.SearchDatastoreSubFolders_Task("[$($ds.Name)]", $spec)) -Server $vc
while ($t.Info.State -in 'queued','running') { Start-Sleep -Milliseconds 300; $t.UpdateViewData('Info') }
$leftFiles = @($t.Info.Result | ForEach-Object { $_.File } | Where-Object { $_ })
if ($leftFiles.Count) { $left += "datastore 檔案 $($leftFiles.Count) 個" }
Check '清除: 無殘留' ($left.Count -eq 0) $(if ($left.Count) { $left -join ', ' } else { 'clean' })
Disconnect-VIServer -Server $vc -Confirm:$false | Out-Null

# ============ 結果 ============
$fail = @($script:Checks | Where-Object { -not $_.Ok })
Write-Host "`n================ 測試結果 ================"
Write-Host ("  通過 {0} / {1}" -f (@($script:Checks).Count - $fail.Count), @($script:Checks).Count)
if ($fail.Count) { $fail | ForEach-Object { Write-Host "  [FAIL] $($_.Label) -> $($_.Got)" -ForegroundColor Red }; exit 1 }
Write-Host "  ALL PASS" -ForegroundColor Green
exit 0
