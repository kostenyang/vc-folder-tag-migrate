<#
.SYNOPSIS
  端到端自我測試：在來源 vC 建一組帶前綴的測試物件 → 匯出 CSV → 匯入目標 vC
  → 用獨立查詢驗證 → 兩邊清乾淨。全部通過回傳 exit code 0。

.EXAMPLE
  ./Test-VcMeta.ps1 -SourceServer 10.0.0.101 -SourcePassword 'xxx' -SourceDatacenter Datacenter `
                    -TargetServer 10.0.1.19  -TargetPassword 'yyy' -TargetDatacenter m01-dc01 `
                    -NotesTestVM vcda-testvm01

.NOTES
  - 只會動 -Prefix（預設 zz-migtest）開頭的物件，以及 -NotesTestVM 那台 VM 的 Notes（測完還原）。
  - -KeepTestObjects 可保留測試物件不清除，方便到 UI 上看。
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$SourceServer,
    [string]$SourceUser = 'administrator@vsphere.local',
    [Parameter(Mandatory)][string]$SourcePassword,
    [Parameter(Mandatory)][string]$SourceDatacenter,

    [Parameter(Mandatory)][string]$TargetServer,
    [string]$TargetUser = 'administrator@vsphere.local',
    [Parameter(Mandatory)][string]$TargetPassword,
    [Parameter(Mandatory)][string]$TargetDatacenter,

    [string]$NotesTestVM,                       # 目標端一台可以被寫 Notes 的測試 VM（測完清空）
    [string]$Prefix = 'zz-migtest',
    [string]$WorkDir = "$env:TEMP\vc-meta-selftest",
    [switch]$KeepTestObjects
)

$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding $false } catch { }
Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -DisplayDeprecationWarnings $false -Scope Session -Confirm:$false | Out-Null

$here      = Split-Path -Parent $MyInvocation.MyCommand.Path
$folderTop = "$Prefix-folder"
$folderSub = 'Level2'
$catName   = $Prefix
$tagName   = "$Prefix-tag-a"
$caName    = "$Prefix-attr"
$caValue   = '測試值-中文-2026'
$noteText  = "自我測試備註 $Prefix $(Get-Date -Format yyyy-MM-dd)"

$script:Checks = New-Object System.Collections.ArrayList
function Check {
    param([string]$Label, [bool]$Ok, [string]$Got)
    [void]$script:Checks.Add([pscustomobject]@{ Label = $Label; Ok = $Ok; Got = $Got })
    Write-Host ("  {0} {1,-40} {2}" -f $(if ($Ok) { '[PASS]' } else { '[FAIL]' }), $Label, $Got)
}
function Connect-Vc { param($s, $u, $p) Connect-VIServer -Server $s -User $u -Password $p }

# ============ 1. 來源端建立測試物件 ============
Write-Host "`n=== [1/6] 在來源 $SourceServer 建立測試物件 ==="
$src = Connect-Vc $SourceServer $SourceUser $SourcePassword
$sdc  = Get-Datacenter -Name $SourceDatacenter -Server $src
$sroot = Get-Folder -Id $sdc.ExtensionData.VmFolder.ToString() -Server $src
$sl1 = Get-Folder -Location $sroot -Name $folderTop -NoRecursion -Server $src -ErrorAction SilentlyContinue
if (-not $sl1) { $sl1 = New-Folder -Name $folderTop -Location $sroot -Server $src }
$sl2 = Get-Folder -Location $sl1 -Name $folderSub -NoRecursion -Server $src -ErrorAction SilentlyContinue
if (-not $sl2) { $sl2 = New-Folder -Name $folderSub -Location $sl1 -Server $src }
$scat = Get-TagCategory -Name $catName -Server $src -ErrorAction SilentlyContinue
if (-not $scat) { $scat = New-TagCategory -Name $catName -Description '搬遷測試分類' -Cardinality Single -EntityType All -Server $src }
$stag = Get-Tag -Name $tagName -Category $scat -Server $src -ErrorAction SilentlyContinue
if (-not $stag) { $stag = New-Tag -Name $tagName -Category $scat -Description '測試標籤 A' -Server $src }
if (-not (Get-TagAssignment -Entity $sl2 -Tag $stag -Server $src -ErrorAction SilentlyContinue)) {
    $null = New-TagAssignment -Tag $stag -Entity $sl2 -Server $src -Confirm:$false
}
if (-not (Get-CustomAttribute -Name $caName -Server $src -ErrorAction SilentlyContinue)) {
    $null = New-CustomAttribute -Name $caName -Server $src
}
$null = $sl2 | Set-Annotation -CustomAttribute $caName -Value $caValue -Server $src
Write-Host "  建立: $folderTop/$folderSub, $catName/$tagName, $caName=$caValue"
Disconnect-VIServer -Server $src -Confirm:$false | Out-Null

# ============ 2. 匯出 ============
Write-Host "`n=== [2/6] 匯出 ==="
if (Test-Path $WorkDir) { Remove-Item $WorkDir -Recurse -Force }
& "$here\Export-VcMeta.ps1" -Server $SourceServer -User $SourceUser -Password $SourcePassword -OutDir $WorkDir | Out-Null
$exported = @{
    folder = (Import-Csv "$WorkDir\folders.csv"                | Where-Object { $_.Path -eq "$folderTop/$folderSub" })
    cat    = (Import-Csv "$WorkDir\tag-categories.csv"         | Where-Object { $_.Name -eq $catName })
    tag    = (Import-Csv "$WorkDir\tags.csv"                   | Where-Object { $_.Name -eq $tagName })
    assign = (Import-Csv "$WorkDir\tag-assignments.csv"        | Where-Object { $_.Tag  -eq $tagName })
    ca     = (Import-Csv "$WorkDir\custom-attributes.csv"      | Where-Object { $_.Name -eq $caName })
    caval  = (Import-Csv "$WorkDir\custom-attribute-values.csv"| Where-Object { $_.AttributeName -eq $caName })
}
Check '匯出: 兩層資料夾'        ($null -ne $exported.folder) "$folderTop/$folderSub"
Check '匯出: tag 分類/標籤'     (($null -ne $exported.cat) -and ($null -ne $exported.tag)) "$catName/$tagName"
Check '匯出: 資料夾上的標籤指派' ($null -ne $exported.assign) "$($exported.assign.EntityType):$($exported.assign.EntityPath)"
Check '匯出: 自訂屬性 + 中文值'  (($null -ne $exported.ca) -and ($exported.caval.Value -eq $caValue)) "$caName=$($exported.caval.Value)"

# 測「手動編 CSV 再匯入」：加一列 Notes，故意不填 UUID（走名稱比對）
if ($NotesTestVM) {
    Add-Content -Path "$WorkDir\notes.csv" -Encoding utf8 `
        -Value ('"{0}","VirtualMachine","{1}","","{2}"' -f $SourceDatacenter, $NotesTestVM, $noteText)
    Write-Host "  手動加一列 Notes -> $NotesTestVM（UUID 留空，測名稱 fallback）"
}

# ============ 2b. 只匯出這個資料夾（-Folder 範圍）============
Write-Host "`n=== [2b] 依資料夾範圍匯出 (-Folder $folderTop) ==="
$scopeDir = "$WorkDir-scope"
if (Test-Path $scopeDir) { Remove-Item $scopeDir -Recurse -Force }
& "$here\Export-VcMeta.ps1" -Server $SourceServer -User $SourceUser -Password $SourcePassword -OutDir $scopeDir -Folder $folderTop | Out-Null
$sFolders = @(Import-Csv "$scopeDir\folders.csv")
$sTags    = @(Import-Csv "$scopeDir\tags.csv")
$sCats    = @(Import-Csv "$scopeDir\tag-categories.csv")
$sAttrs   = @(Import-Csv "$scopeDir\custom-attributes.csv")
$sAssign  = @(Import-Csv "$scopeDir\tag-assignments.csv")
$sNotes   = @(Import-Csv "$scopeDir\notes.csv")
Check '範圍匯出: 只有這棵子樹的資料夾' `
    (($sFolders.Count -eq 2) -and (@($sFolders | Where-Object { $_.Path -notlike "$folderTop*" }).Count -eq 0)) `
    ("$($sFolders.Count) 筆: " + (($sFolders.Path) -join ', '))
Check '範圍匯出: 只帶到有用到的分類/標籤' `
    (($sCats.Count -eq 1) -and ($sTags.Count -eq 1) -and ($sTags[0].Name -eq $tagName)) `
    ("分類 $($sCats.Count) / 標籤 $($sTags.Count)")
Check '範圍匯出: 只帶到有用到的屬性定義' `
    (($sAttrs.Count -eq 1) -and ($sAttrs[0].Name -eq $caName)) ("$($sAttrs.Count) 筆")
Check '範圍匯出: 指派/Notes 不含範圍外物件' `
    ((@($sAssign | Where-Object { $_.EntityPath -notlike "$folderTop*" }).Count -eq 0) -and ($sNotes.Count -eq 0)) `
    ("指派 $($sAssign.Count) / notes $($sNotes.Count)")

# ============ 3. DryRun ============
Write-Host "`n=== [3/6] DryRun 匯入 $TargetServer ==="
$dryReport = "$WorkDir\dryrun.csv"
& "$here\Import-VcMeta.ps1" -Server $TargetServer -User $TargetUser -Password $TargetPassword -InDir $WorkDir `
    -DatacenterMap "$SourceDatacenter=$TargetDatacenter" -Include Folders,Tags,CustomAttributes,Notes `
    -DryRun -ReportPath $dryReport | Out-Null
$dry = Import-Csv $dryReport
$wouldFolders = @($dry | Where-Object { $_.Action -eq 'WouldCreate' -and $_.Section -eq 'Folder' -and $_.Target -like "*$folderTop*" })
Check 'DryRun: 父/子資料夾各列一次' ($wouldFolders.Count -eq 2) (($wouldFolders.Target | ForEach-Object { ($_ -split '\|')[-1] }) -join ', ')
Check 'DryRun: 不會誤動既有物件'    (@($dry | Where-Object { $_.Action -in 'Created','Set','Assigned' }).Count -eq 0) '無實際寫入'

# ============ 3b. 匯入端的資料夾範圍過濾 ============
# 用「整台 vC 的完整匯出」但加 -Folder，應該只會動到這棵子樹；
# 手動加的那筆 Notes 指向範圍外的 VM，必須被濾掉。
Write-Host "`n=== [3b] DryRun + -Folder（匯入端過濾） ==="
$dryScope = "$WorkDir\dryrun-scope.csv"
& "$here\Import-VcMeta.ps1" -Server $TargetServer -User $TargetUser -Password $TargetPassword -InDir $WorkDir `
    -DatacenterMap "$SourceDatacenter=$TargetDatacenter" -Include Folders,Tags,CustomAttributes,Notes `
    -Folder $folderTop -DryRun -ReportPath $dryScope | Out-Null
$ds = Import-Csv $dryScope
$dsWould = @($ds | Where-Object { $_.Action -like 'Would*' })
$outOfScope = @($dsWould | Where-Object {
    $_.Target -notlike "*$folderTop*" -and $_.Target -notlike "*$Prefix*"
})
Check '範圍匯入: 不碰範圍外的東西' ($outOfScope.Count -eq 0) `
    ("Would* 共 $($dsWould.Count) 筆，範圍外 $($outOfScope.Count) 筆")
Check '範圍匯入: 濾掉範圍外 VM 的 Notes' `
    (@($ds | Where-Object { $_.Section -eq 'Notes' }).Count -eq 0) `
    ("Notes 相關 $((@($ds | Where-Object { $_.Section -eq 'Notes' })).Count) 筆")

# ============ 4. 正式匯入 ============
Write-Host "`n=== [4/6] 正式匯入 ==="
$applyReport = "$WorkDir\apply.csv"
& "$here\Import-VcMeta.ps1" -Server $TargetServer -User $TargetUser -Password $TargetPassword -InDir $WorkDir `
    -DatacenterMap "$SourceDatacenter=$TargetDatacenter" -Include Folders,Tags,CustomAttributes,Notes `
    -ReportPath $applyReport | Out-Null
$apply = Import-Csv $applyReport
Check '匯入: 有實際建立動作' (@($apply | Where-Object { $_.Action -in 'Created','Set','Assigned' }).Count -ge 5) `
    ("Created/Set/Assigned = " + @($apply | Where-Object { $_.Action -in 'Created','Set','Assigned' }).Count)
Check '匯入: 無 Failed'      (@($apply | Where-Object { $_.Action -eq 'Failed' }).Count -eq 0) `
    ("Failed = " + @($apply | Where-Object { $_.Action -eq 'Failed' }).Count)

# ============ 5. 獨立查詢驗證目標端 ============
Write-Host "`n=== [5/6] 目標端獨立驗證 ==="
$tgt = Connect-Vc $TargetServer $TargetUser $TargetPassword
$tdc   = Get-Datacenter -Name $TargetDatacenter -Server $tgt
$troot = Get-Folder -Id $tdc.ExtensionData.VmFolder.ToString() -Server $tgt
$tl1 = Get-Folder -Location $troot -Name $folderTop -NoRecursion -Server $tgt -ErrorAction SilentlyContinue
$tl2 = if ($tl1) { Get-Folder -Location $tl1 -Name $folderSub -NoRecursion -Server $tgt -ErrorAction SilentlyContinue } else { $null }
Check '目標: 資料夾樹(含父子關係)' ($null -ne $tl2) $(if ($tl2) { "$($tl1.Name)/$($tl2.Name) parent=$($tl2.Parent.Name)" } else { '找不到' })

$tcat = Get-TagCategory -Name $catName -Server $tgt -ErrorAction SilentlyContinue
Check '目標: tag 分類屬性一致' (($null -ne $tcat) -and ($tcat.Cardinality -eq 'Single')) `
    $(if ($tcat) { "cardinality=$($tcat.Cardinality) entityType=$($tcat.EntityType -join ',')" } else { '找不到' })

$ttag = if ($tcat) { Get-Tag -Name $tagName -Category $tcat -Server $tgt -ErrorAction SilentlyContinue } else { $null }
$tasg = if ($tl2 -and $ttag) { Get-TagAssignment -Entity $tl2 -Tag $ttag -Server $tgt -ErrorAction SilentlyContinue } else { $null }
Check '目標: 標籤貼在正確路徑的資料夾' ($null -ne $tasg) $(if ($tasg) { "$($tasg.Entity.Name) <- $($tasg.Tag.Name)" } else { '沒有指派' })

$tval = if ($tl2) { ($tl2 | Get-Annotation -CustomAttribute $caName -Server $tgt -ErrorAction SilentlyContinue).Value } else { $null }
Check '目標: 自訂屬性中文值' ($tval -eq $caValue) "值 = '$tval'"

if ($NotesTestVM) {
    $tvm = Get-VM -Name $NotesTestVM -Server $tgt -ErrorAction SilentlyContinue
    Check '目標: VM Notes(名稱 fallback)' ($tvm -and $tvm.Notes -eq $noteText) "Notes = '$($tvm.Notes)'"
}
Disconnect-VIServer -Server $tgt -Confirm:$false | Out-Null

# ============ 6. 清除 ============
if ($KeepTestObjects) {
    Write-Host "`n=== [6/6] -KeepTestObjects 指定，保留測試物件 ==="
} else {
    Write-Host "`n=== [6/6] 清除兩邊測試物件 ==="
    foreach ($t in @(
        @{ s = $TargetServer; u = $TargetUser; p = $TargetPassword; dc = $TargetDatacenter },
        @{ s = $SourceServer; u = $SourceUser; p = $SourcePassword; dc = $SourceDatacenter }
    )) {
        $c = Connect-Vc $t.s $t.u $t.p
        try {
            if ($NotesTestVM) {
                $vm = Get-VM -Name $NotesTestVM -Server $c -ErrorAction SilentlyContinue
                if ($vm -and $vm.Notes -like "*$Prefix*") { $null = Set-VM -VM $vm -Notes '' -Server $c -Confirm:$false }
            }
            $cc = Get-TagCategory -Name $catName -Server $c -ErrorAction SilentlyContinue
            if ($cc) { Remove-TagCategory -Category $cc -Server $c -Confirm:$false }     # 連標籤與指派一起移除
            $ca = Get-CustomAttribute -Name $caName -Server $c -ErrorAction SilentlyContinue
            if ($ca) { Remove-CustomAttribute -CustomAttribute $ca -Server $c -Confirm:$false }
            $dcv  = Get-Datacenter -Name $t.dc -Server $c
            $rt   = Get-Folder -Id $dcv.ExtensionData.VmFolder.ToString() -Server $c
            $top  = Get-Folder -Location $rt -Name $folderTop -NoRecursion -Server $c -ErrorAction SilentlyContinue
            if ($top) {
                $inside = Get-VM -Location $top -Server $c -ErrorAction SilentlyContinue
                if ($inside) { Write-Host "  !! $($t.s) 的 $folderTop 內有 VM，保留不刪" }
                else { Remove-Folder -Folder $top -DeletePermanently -Server $c -Confirm:$false }
            }
            $left = @()
            if (Get-TagCategory -Name $catName -Server $c -ErrorAction SilentlyContinue)   { $left += 'tagCategory' }
            if (Get-CustomAttribute -Name $caName -Server $c -ErrorAction SilentlyContinue) { $left += 'customAttribute' }
            if (Get-Folder -Name $folderTop -Server $c -ErrorAction SilentlyContinue)       { $left += 'folder' }
            Check "清除: $($t.s) 無殘留" ($left.Count -eq 0) $(if ($left.Count) { $left -join ', ' } else { 'clean' })
        } finally { Disconnect-VIServer -Server $c -Confirm:$false | Out-Null }
    }
}

# ============ 結果 ============
$fail = @($script:Checks | Where-Object { -not $_.Ok })
Write-Host "`n================ 測試結果 ================"
Write-Host ("  通過 {0} / {1}" -f (@($script:Checks).Count - $fail.Count), @($script:Checks).Count)
if ($fail.Count) {
    $fail | ForEach-Object { Write-Host "  [FAIL] $($_.Label) -> $($_.Got)" -ForegroundColor Red }
    exit 1
}
Write-Host "  ALL PASS" -ForegroundColor Green
exit 0
