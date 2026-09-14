<#
.SYNOPSIS
  只搬資料夾樹：把 vCenter A 的 VM 資料夾結構整個建到 vCenter B（可選連 VM 一起放進去）。

.EXAMPLE
  ./Copy-VcFolders.ps1 -SourceServer vcA -SourcePassword 'x' -TargetServer vcB -TargetPassword 'y' -DatacenterMap 'DC-A=DC-B' -DryRun
  ./Copy-VcFolders.ps1 ... -MoveVMs          # 目標端已有同名/同 UUID 的 VM 就搬進對應資料夾
  ./Copy-VcFolders.ps1 ... -Folder 'Linux'   # 只建這棵子樹
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
    [switch]$MoveVMs,
    [switch]$DryRun,
    [string]$WorkDir
)
$p = @{} + $PSBoundParameters
$p.Include = if ($MoveVMs) { @('Folders','VMPlacement') } else { @('Folders') }
& (Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) 'Copy-VcMeta.ps1') @p
