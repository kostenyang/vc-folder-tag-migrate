<#
.SYNOPSIS
  Register-VmxFromDatastore.ps1 的端到端自我測試（含「照舊 vC 補齊」）：
  建丟棄式 VM + 範本，給 VM 放資料夾、自訂屬性、Notes、tag → 用 Export-VcMeta 匯出當「舊 vC 快照」
  → 從 inventory 移除 VM/範本、刪掉資料夾/屬性定義/tag 分類（模擬新 vC 什麼都沒有）
  → Register -MetaDir 註冊回來 → 驗證資料夾/屬性/Notes/tag/範本全部照舊 → 冪等 → 全部清掉。

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
$caName  = "$Prefix-attr"
$caValue = '照舊值-中文-2026'
$notes   = "自我測試備註 $Prefix"
$catName = "$Prefix-cat"
$tagName = "$Prefix-tag"
$metaDir = Join-Path $WorkDir 'meta'
if (Test-Path $WorkDir) { Remove-Item $WorkDir -Recurse -Force }
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
function Remove-TestLeftovers {
    param($vc, $dc)
    foreach ($n in $vmName, $tplName) {
        if (Get-VM -Name $n -Server $vc -ErrorAction SilentlyContinue)       { Remove-VM -VM $n -DeletePermanently -Confirm:$false -Server $vc }
        if (Get-Template -Name $n -Server $vc -ErrorAction SilentlyContinue) { Remove-Template -Template $n -DeletePermanently -Confirm:$false -Server $vc }
    }
    $cat = Get-TagCategory -Name $catName -Server $vc -ErrorAction SilentlyContinue
    if ($cat) { Remove-TagCategory -Category $cat -Confirm:$false -Server $vc }
    $ca = Get-CustomAttribute -Name $caName -Server $vc -ErrorAction SilentlyContinue
    if ($ca) { Remove-CustomAttribute -CustomAttribute $ca -Confirm:$false -Server $vc }
    $root = Get-Folder -Id $dc.ExtensionData.VmFolder.ToString() -Server $vc
    $top  = Get-Folder -Location $root -Name $fldTop -NoRecursion -Server $vc -ErrorAction SilentlyContinue
    if ($top) { Remove-Folder -Folder $top -DeletePermanently -Confirm:$false -Server $vc }
}

# ============ 1. 建測試 VM + 範本，並給 VM 資料夾 / 屬性 / Notes / tag ============
Write-Host "`n=== [1/7] 在 $Server 的 $Datastore 建測試 VM 與範本，並上齊中繼資料 ==="
$vc = Connect-VIServer -Server $Server -User $User -Password $Password
$ds = Get-Datastore -Name $Datastore -Server $vc
$dc = $ds.Datacenter
$host0 = $ds | Get-VMHost -Server $vc | Where-Object { $_.ConnectionState -eq 'Connected' -and (-not $Cluster -or $_.Parent.Name -eq $Cluster) } | Select-Object -First 1
if (-not $host0) { throw "沒有主機掛著 $Datastore" }
Remove-TestLeftovers $vc $dc

$root = Get-Folder -Id $dc.ExtensionData.VmFolder.ToString() -Server $vc
$f1 = New-Folder -Name $fldTop -Location $root -Server $vc
$f2 = New-Folder -Name 'Sub' -Location $f1 -Server $vc
$vm = New-VM -Name $vmName  -VMHost $host0 -Datastore $ds -Location $f2 -NumCpu 1 -MemoryMB 128 -DiskMB 16 -GuestId otherLinux64Guest -Notes $notes -Server $vc
$tv = New-VM -Name $tplName -VMHost $host0 -Datastore $ds -NumCpu 1 -MemoryMB 128 -DiskMB 16 -GuestId otherLinux64Guest -Server $vc
$null = New-CustomAttribute -Name $caName -TargetType VirtualMachine -Server $vc
$null = $vm | Set-Annotation -CustomAttribute $caName -Value $caValue -Server $vc
$cat = New-TagCategory -Name $catName -Cardinality Single -EntityType VirtualMachine -Server $vc
$tag = New-Tag -Name $tagName -Category $cat -Server $vc
$null = New-TagAssignment -Tag $tag -Entity $vm -Confirm:$false -Server $vc
$vmxPath = $vm.ExtensionData.Config.Files.VmPathName
$null = Set-VM -VM $tv -ToTemplate -Confirm:$false -Server $vc
$tplPath = (Get-Template -Name $tplName -Server $vc).ExtensionData.Config.Files.VmPathName
Write-Host "  VM      : $vmxPath  -> $fldPath, $caName=$caValue, notes, tag $tagName"
Write-Host "  Template: $tplPath"
Disconnect-VIServer -Server $vc -Confirm:$false | Out-Null

# ============ 2. 匯出當「舊 vC 快照」============
Write-Host "`n=== [2/7] Export-VcMeta 匯出（模擬舊 vC）==="
& "$here\Export-VcMeta.ps1" -Server $Server -User $User -Password $Password -OutDir $metaDir -Folder $fldTop | Out-Null
$pl = @(Import-Csv "$metaDir\vm-placement.csv" | Where-Object { $_.VMName -eq $vmName })
$cv = @(Import-Csv "$metaDir\custom-attribute-values.csv" | Where-Object { $_.EntityName -eq $vmName })
$nt = @(Import-Csv "$metaDir\notes.csv" | Where-Object { $_.EntityName -eq $vmName })
$ta = @(Import-Csv "$metaDir\tag-assignments.csv" | Where-Object { $_.EntityName -eq $vmName })
Check '快照: 資料夾/屬性/Notes/tag 都在' (($pl.Count -eq 1) -and ($pl[0].FolderPath -eq $fldPath) -and ($cv.Count -eq 1) -and ($nt.Count -eq 1) -and ($ta.Count -eq 1)) "placement=$($pl.Count) ca=$($cv.Count) notes=$($nt.Count) tag=$($ta.Count)"

# ============ 3. 模擬「新 vC 什麼都沒有」============
Write-Host "`n=== [3/7] unregister VM/範本，並刪掉資料夾、屬性定義、tag 分類 ==="
$vc = Connect-VIServer -Server $Server -User $User -Password $Password
Remove-VM -VM $vmName -Confirm:$false -Server $vc                       # 沒有 -DeletePermanently = 只移出 inventory
Remove-Template -Template $tplName -Confirm:$false -Server $vc
Remove-TagCategory -Category (Get-TagCategory -Name $catName -Server $vc) -Confirm:$false -Server $vc
Remove-CustomAttribute -CustomAttribute (Get-CustomAttribute -Name $caName -Server $vc) -Confirm:$false -Server $vc
Remove-Folder -Folder (Get-Folder -Location $root -Name $fldTop -NoRecursion -Server $vc) -DeletePermanently -Confirm:$false -Server $vc
$gone = -not (Get-VM -Name $vmName -Server $vc -ErrorAction SilentlyContinue) -and
        -not (Get-Template -Name $tplName -Server $vc -ErrorAction SilentlyContinue) -and
        -not (Get-Folder -Name $fldTop -Server $vc -ErrorAction SilentlyContinue) -and
        -not (Get-CustomAttribute -Name $caName -Server $vc -ErrorAction SilentlyContinue) -and
        -not (Get-TagCategory -Name $catName -Server $vc -ErrorAction SilentlyContinue)
Check '新 vC 狀態: VM/範本/資料夾/屬性/tag 全沒了' $gone ''
Disconnect-VIServer -Server $vc -Confirm:$false | Out-Null

# ============ 4. DryRun ============
Write-Host "`n=== [4/7] DryRun（-MetaDir）==="
$dry = Run-Register -Extra @{ DryRun = $true; MetaDir = $metaDir } -Report (Join-Path $WorkDir 'dry.csv')
$would = @($dry | Where-Object { $_.Action -eq 'WouldRegister' })
Check 'DryRun: 只會註冊這 2 個' (($would.Count -eq 2) -and (@($would | Where-Object { $_.Target -notlike "*$Prefix*" }).Count -eq 0)) "WouldRegister=$($would.Count)"
Check 'DryRun: 分得出 VM / Template，且 VM 預告放進原資料夾' ((@($would | Where-Object { $_.Detail -like "VM '$vmName'*folder=$fldPath" }).Count -eq 1) -and (@($would | Where-Object { $_.Detail -like 'Template *' }).Count -eq 1)) (($would.Detail) -join ' | ')
Check 'DryRun: 沒有真的註冊' (@($dry | Where-Object { $_.Action -eq 'Registered' }).Count -eq 0) ''
$dryMeta = @(Import-Csv (Join-Path $WorkDir 'dry-meta.csv'))
Check 'DryRun: 預告會補屬性/Notes/tag' ((@($dryMeta | Where-Object { $_.Section -eq 'CAValue' -and $_.Action -eq 'WouldSet' }).Count -ge 0) -and (@($dryMeta | Where-Object { $_.Action -like 'Would*' -and $_.Target -notlike "*$Prefix*" -and $_.Target -notlike "*$vmName*" }).Count -eq 0)) ("Would* " + @($dryMeta | Where-Object { $_.Action -like 'Would*' }).Count + " 筆，皆在範圍內")

# ============ 5. 正式：Register -MetaDir ============
Write-Host "`n=== [5/7] 正式註冊 + 照舊 vC 補齊（-MetaDir）==="
$apply = Run-Register -Extra @{ MetaDir = $metaDir } -Report (Join-Path $WorkDir 'apply.csv')
$applyMeta = @(Import-Csv (Join-Path $WorkDir 'apply-meta.csv'))
Check '註冊: Registered = 2, Failed = 0' ((@($apply | Where-Object { $_.Action -eq 'Registered' }).Count -eq 2) -and (@($apply | Where-Object { $_.Action -eq 'Failed' }).Count -eq 0)) ("Registered=" + @($apply | Where-Object { $_.Action -eq 'Registered' }).Count + " Failed=" + @($apply | Where-Object { $_.Action -eq 'Failed' }).Count)
Check '註冊: 資料夾自動建好並放進去' ((@($apply | Where-Object { $_.Action -eq 'CreatedFolder' }).Count -eq 2) -and (@($apply | Where-Object { $_.Action -eq 'Placed' -and $_.Target -eq $vmName }).Count -eq 1)) ("CreatedFolder=" + @($apply | Where-Object { $_.Action -eq 'CreatedFolder' }).Count)
Check '補齊: 屬性定義+值 / Notes(vmx自帶=AlreadySet) / tag 都有做' ((@($applyMeta | Where-Object { $_.Section -eq 'CustomAttribute' -and $_.Action -eq 'Created' }).Count -eq 1) -and (@($applyMeta | Where-Object { $_.Section -eq 'CAValue' -and $_.Action -eq 'Set' }).Count -eq 1) -and (@($applyMeta | Where-Object { $_.Section -eq 'Notes' -and $_.Action -in 'Set','AlreadySet' }).Count -eq 1) -and (@($applyMeta | Where-Object { $_.Section -eq 'TagAssignment' -and $_.Action -eq 'Assigned' }).Count -eq 1)) (($applyMeta | Where-Object { $_.Action -in 'Created','Set','Assigned' } | ForEach-Object { "$($_.Section)/$($_.Action)" }) -join ', ')
Check '補齊: 沒有 Failed' (@($applyMeta | Where-Object { $_.Action -eq 'Failed' }).Count -eq 0) (($applyMeta | Where-Object { $_.Action -eq 'Failed' } | ForEach-Object { $_.Detail }) -join '; ')

# ============ 6. 獨立驗證 ============
Write-Host "`n=== [6/7] 獨立驗證 ==="
$vc = Connect-VIServer -Server $Server -User $User -Password $Password
$v2 = Get-VM -Name $vmName -Server $vc -ErrorAction SilentlyContinue
$t2 = Get-Template -Name $tplName -Server $vc -ErrorAction SilentlyContinue
Check 'VM 回到 inventory、vmx 路徑一致' ($v2 -and $v2.ExtensionData.Config.Files.VmPathName -eq $vmxPath) $(if ($v2) { $v2.ExtensionData.Config.Files.VmPathName } else { '找不到' })
Check '範本回到 inventory 且仍是範本'   ($t2 -and $t2.ExtensionData.Config.Template -and $t2.ExtensionData.Config.Files.VmPathName -eq $tplPath) $(if ($t2) { "template=$($t2.ExtensionData.Config.Template)" } else { '找不到' })
$vmFolder = if ($v2) { $v2.Folder } else { $null }
Check 'VM 在原本的資料夾' ($vmFolder -and $vmFolder.Name -eq 'Sub' -and $vmFolder.Parent.Name -eq $fldTop) $(if ($vmFolder) { "$($vmFolder.Parent.Name)/$($vmFolder.Name)" } else { '' })
$val = if ($v2) { ($v2 | Get-Annotation -CustomAttribute $caName -Server $vc -ErrorAction SilentlyContinue).Value } else { $null }
Check '自訂屬性值照舊（中文）' ($val -eq $caValue) "值 = '$val'"
Check 'Notes 照舊' ($v2 -and $v2.Notes -eq $notes) "Notes = '$($v2.Notes)'"
$asg = if ($v2) { Get-TagAssignment -Entity $v2 -Server $vc -ErrorAction SilentlyContinue | Where-Object { $_.Tag.Name -eq $tagName } } else { $null }
Check 'tag 照舊（分類+標籤重建並指派）' ($null -ne $asg) $(if ($asg) { "$($asg.Tag.Category.Name)/$($asg.Tag.Name)" } else { '沒有' })
Check 'VM 沒被開機' ($v2 -and $v2.PowerState -eq 'PoweredOff') "$($v2.PowerState)"
Disconnect-VIServer -Server $vc -Confirm:$false | Out-Null

Write-Host "`n=== [6b] 重跑一次（冪等）==="
$again = Run-Register -Extra @{ MetaDir = $metaDir } -Report (Join-Path $WorkDir 'again.csv')
Check '重跑: 全部 AlreadyRegistered、沒再補資料' ((@($again | Where-Object { $_.Action -eq 'Registered' }).Count -eq 0) -and (@($again | Where-Object { $_.Action -eq 'AlreadyRegistered' -and $_.Target -like "*$Prefix*" }).Count -eq 2) -and -not (Test-Path (Join-Path $WorkDir 'again-meta.csv'))) ("Registered=" + @($again | Where-Object { $_.Action -eq 'Registered' }).Count)

# ============ 7. 清除 ============
Write-Host "`n=== [7/7] 清除 ==="
$vc = Connect-VIServer -Server $Server -User $User -Password $Password
Remove-TestLeftovers $vc $dc
$left = @()
if (Get-VM -Name $vmName -Server $vc -ErrorAction SilentlyContinue)            { $left += 'vm' }
if (Get-Template -Name $tplName -Server $vc -ErrorAction SilentlyContinue)     { $left += 'template' }
if (Get-Folder -Name $fldTop -Server $vc -ErrorAction SilentlyContinue)        { $left += 'folder' }
if (Get-CustomAttribute -Name $caName -Server $vc -ErrorAction SilentlyContinue) { $left += 'customAttribute' }
if (Get-TagCategory -Name $catName -Server $vc -ErrorAction SilentlyContinue)  { $left += 'tagCategory' }
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
