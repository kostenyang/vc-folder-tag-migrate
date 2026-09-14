<#
.SYNOPSIS
  一支腳本直接 vCenter A → vCenter B：資料夾 / Tag / 自訂屬性 / Notes / VM 位置。
  內部就是 Export-VcMeta → Import-VcMeta，CSV 會留在 -WorkDir 給你看或改。

.EXAMPLE
  # 只搬自訂屬性（定義 + 值）
  ./Copy-VcMeta.ps1 -SourceServer vcA -SourcePassword 'x' -TargetServer vcB -TargetPassword 'y' -Include CustomAttributes

  # 資料夾樹整個建到新 vC，連 VM 也搬進去
  ./Copy-VcMeta.ps1 -SourceServer vcA -SourcePassword 'x' -TargetServer vcB -TargetPassword 'y' -Include Folders,VMPlacement -MoveVMs -DatacenterMap 'DC-A=DC-B'

  # 只搬某個資料夾子樹的一切
  ./Copy-VcMeta.ps1 ... -Folder 'Linux'
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$SourceServer,
    [string]$SourceUser = 'administrator@vsphere.local',
    [Parameter(Mandatory)][string]$SourcePassword,

    [Parameter(Mandatory)][string]$TargetServer,
    [string]$TargetUser = 'administrator@vsphere.local',
    [Parameter(Mandatory)][string]$TargetPassword,

    [ValidateSet('Folders','VMPlacement','Tags','CustomAttributes','Notes')]
    [string[]]$Include = @('Folders','Tags','CustomAttributes','Notes'),
    [string[]]$Folder,                 # 只搬這些 VM 資料夾子樹
    [string[]]$Datacenter,             # 只搬這些來源 Datacenter
    [string[]]$DatacenterMap,          # 'DC-A=DC-B'
    [switch]$AllDefinitions,
    [switch]$MoveVMs,
    [switch]$DryRun,
    [string]$WorkDir = (Join-Path (Get-Location) ('copy-{0}-{1:yyyyMMdd-HHmmss}' -f ($SourceServer -replace '[^\w.-]','_'), (Get-Date)))
)

$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding $false } catch { }
$here = Split-Path -Parent $MyInvocation.MyCommand.Path

# 匯出：-Folder 範圍時需要 VMPlacement 來判定哪些 VM 在範圍內
$exportInclude = @($Include)
if ($Folder -and $exportInclude -notcontains 'VMPlacement') { $exportInclude += 'VMPlacement' }

Write-Host "==================== [1/2] 匯出 $SourceServer ===================="
$exp = @{ Server = $SourceServer; User = $SourceUser; Password = $SourcePassword; OutDir = $WorkDir; Include = $exportInclude }
if ($Folder)         { $exp.Folder = $Folder }
if ($Datacenter)     { $exp.Datacenter = $Datacenter }
if ($AllDefinitions) { $exp.AllDefinitions = $true }
& "$here\Export-VcMeta.ps1" @exp

Write-Host "`n==================== [2/2] 匯入 $TargetServer ===================="
$imp = @{ Server = $TargetServer; User = $TargetUser; Password = $TargetPassword; InDir = $WorkDir; Include = $Include }
if ($Folder)        { $imp.Folder = $Folder }
if ($DatacenterMap) { $imp.DatacenterMap = $DatacenterMap }
if ($MoveVMs)       { $imp.MoveVMs = $true }
if ($DryRun)        { $imp.DryRun = $true }
& "$here\Import-VcMeta.ps1" @imp

Write-Host "`n[+] CSV 與報告都在 $WorkDir"
