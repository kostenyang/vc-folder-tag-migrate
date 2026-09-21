#Requires -Version 5.1
<#
.SYNOPSIS
  Folder tree only: rebuild the VM folder structure of vCenter A on vCenter B (optionally move the VMs into it).

.EXAMPLE
  .\Copy-VcFolders.ps1 -SourceServer vcA -SourcePassword 'x' -TargetServer vcB -TargetPassword 'y' -DatacenterMap 'DC-A=DC-B' -DryRun
  .\Copy-VcFolders.ps1 ... -MoveVMs          # VMs that already exist on the target (same name / UUID) are moved into their folders
  .\Copy-VcFolders.ps1 ... -Folder 'Linux'   # only this subtree
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
