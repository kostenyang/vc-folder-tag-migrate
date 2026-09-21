<#
.SYNOPSIS
  Custom attributes only (definitions + the value on every object): vCenter A to vCenter B in one script.

.EXAMPLE
  .\Copy-VcCustomAttributes.ps1 -SourceServer vcA -SourcePassword 'x' -TargetServer vcB -TargetPassword 'y' -DatacenterMap 'DC-A=DC-B' -DryRun
  .\Copy-VcCustomAttributes.ps1 ... -Folder 'Linux'     # only the VMs inside one folder
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
    [switch]$AllDefinitions,           # also create attribute definitions that no object uses
    [switch]$DryRun,
    [string]$WorkDir
)
$p = @{} + $PSBoundParameters
$p.Include = @('CustomAttributes')
& (Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) 'Copy-VcMeta.ps1') @p
