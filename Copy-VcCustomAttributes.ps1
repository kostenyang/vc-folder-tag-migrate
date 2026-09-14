<#
.SYNOPSIS
  只搬自訂屬性（定義 + 每個物件的值）：vCenter A → vCenter B，一支腳本。

.EXAMPLE
  ./Copy-VcCustomAttributes.ps1 -SourceServer vcA -SourcePassword 'x' -TargetServer vcB -TargetPassword 'y' -DatacenterMap 'DC-A=DC-B' -DryRun
  ./Copy-VcCustomAttributes.ps1 ... -Folder 'Linux'     # 只搬某資料夾裡 VM 的屬性
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$SourceServer,
    [string]$SourceUser = 'administrator@vsphere.local',
    [Parameter(Mandatory)][string]$SourcePassword,
    [Parameter(Mandatory)][string]$TargetServer,
    [string]$TargetUser = 'administrator@vsphere.local',
    [Parameter(Mandatory)][string]$TargetPassword,
    [string[]]$Folder,
    [string[]]$Datacenter,
    [string[]]$DatacenterMap,
    [switch]$AllDefinitions,           # 連沒被任何物件用到的屬性定義也建
    [switch]$DryRun,
    [string]$WorkDir
)
$p = @{} + $PSBoundParameters
$p.Include = @('CustomAttributes')
& (Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) 'Copy-VcMeta.ps1') @p
