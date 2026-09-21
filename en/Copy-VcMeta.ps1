<#
.SYNOPSIS
  One script, straight from vCenter A to vCenter B: folders / tags / custom
  attributes / notes / VM placement. Internally runs Export-VcMeta then
  Import-VcMeta; the CSV files stay in -WorkDir for review or editing.

.EXAMPLE
  # Custom attributes only (definitions + values)
  .\Copy-VcMeta.ps1 -SourceServer vcA -SourcePassword 'x' -TargetServer vcB -TargetPassword 'y' -Include CustomAttributes

.EXAMPLE
  # Whole folder tree to the new vCenter, and move the VMs into it
  .\Copy-VcMeta.ps1 -SourceServer vcA -SourcePassword 'x' -TargetServer vcB -TargetPassword 'y' -Include Folders,VMPlacement -MoveVMs -DatacenterMap 'DC-A=DC-B'

.EXAMPLE
  # Everything for one folder subtree only
  .\Copy-VcMeta.ps1 ... -Folder 'Linux'
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
    [string[]]$Folder,                 # only these VM folder subtrees
    [string[]]$Datacenter,             # only these source datacenters
    [string[]]$DatacenterMap,          # 'DC-A=DC-B'
    [switch]$AllDefinitions,
    [switch]$MoveVMs,
    [switch]$DryRun,
    [string]$WorkDir = (Join-Path (Get-Location) ('copy-{0}-{1:yyyyMMdd-HHmmss}' -f ($SourceServer -replace '[^\w.-]','_'), (Get-Date)))
)

$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path

# Export: with -Folder, VMPlacement is needed to know which VMs are in scope
$exportInclude = @($Include)
if ($Folder -and $exportInclude -notcontains 'VMPlacement') { $exportInclude += 'VMPlacement' }

Write-Host "==================== [1/2] Export from $SourceServer ===================="
$exp = @{ Server = $SourceServer; User = $SourceUser; Password = $SourcePassword; OutDir = $WorkDir; Include = $exportInclude }
if ($Folder)         { $exp.Folder = $Folder }
if ($Datacenter)     { $exp.Datacenter = $Datacenter }
if ($AllDefinitions) { $exp.AllDefinitions = $true }
& "$here\Export-VcMeta.ps1" @exp

Write-Host "`n==================== [2/2] Import into $TargetServer ===================="
$imp = @{ Server = $TargetServer; User = $TargetUser; Password = $TargetPassword; InDir = $WorkDir; Include = $Include }
if ($Folder)        { $imp.Folder = $Folder }
if ($DatacenterMap) { $imp.DatacenterMap = $DatacenterMap }
if ($MoveVMs)       { $imp.MoveVMs = $true }
if ($DryRun)        { $imp.DryRun = $true }
& "$here\Import-VcMeta.ps1" @imp

Write-Host "`n[+] CSV files and reports are in $WorkDir"
