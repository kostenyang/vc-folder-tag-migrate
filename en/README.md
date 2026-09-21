# vCenter migration scripts (English / ASCII edition)

Same logic as the scripts in the repository root, with every comment and message
in English. Files are **pure ASCII, UTF-8 with BOM, CRLF**, so they open and run
cleanly in Windows PowerShell 5.1, PowerShell ISE, Notepad and VS Code alike.
Use this edition on customer machines.

Verified: the full runbook below was executed end-to-end with **Windows PowerShell
5.1.20348 + PowerCLI 13.3** (the engine behind ISE) against vCenter 8.0.3 -> 9.1.1
with self-signed certificates. CSV files written by 5.1 are UTF-8 with BOM, so
non-ASCII data (e.g. Chinese attribute values) survives the round trip.

## Self-signed certificates / first run

Every script starts with

```powershell
Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -ParticipateInCeip $false -Scope User
Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Scope Session
```

so self-signed vCenter certificates are accepted and the PowerCLI CEIP question
never blocks a first run. Nothing has to be prepared on the machine beyond
`Install-Module VMware.PowerCLI`.

## Runbook (five independent steps, each re-runnable)

```
Step 0  old vC   Export-VcMeta.ps1                                     -> export-A\   folders / VM placement / tags / attributes / notes
Step 1  new vC   Import-VcMeta.ps1 -Include Folders                    folder tree
Step 2  new vC   Import-VcMeta.ps1 -Include CustomAttributes,Tags      attribute definitions, tag categories/tags
Step 3  old vC   Unregister-VmFromOldVc.ps1                            VMs out of the old inventory -> export-A\unregistered.csv
        ...move the datastore to the new vCenter (any time later)...
Step 4  new vC   Register-VmxFromDatastore.ps1 -UnregisteredCsv        register each VM from the CSV straight into its folder; reconcile
Step 5  new vC   Import-VcMeta.ps1 -Include CustomAttributes,Notes,Tags attribute values / notes / tag assignments
```

```powershell
# Step 0 (old vCenter)
.\Export-VcMeta.ps1 -Server <old-vc> -User administrator@vsphere.local -Password '<pw>' -OutDir .\export-A -Folder 'Linux'

# Step 1 (new vCenter)
.\Import-VcMeta.ps1 -Server <new-vc> -User administrator@vsphere.local -Password '<pw>' -InDir .\export-A -DatacenterMap '<oldDC>=<newDC>' -Include Folders

# Step 2 (new vCenter)
.\Import-VcMeta.ps1 -Server <new-vc> -User administrator@vsphere.local -Password '<pw>' -InDir .\export-A -DatacenterMap '<oldDC>=<newDC>' -Include CustomAttributes,Tags

# Step 3 (old vCenter)  -- filters combine as an intersection
.\Unregister-VmFromOldVc.ps1 -Server <old-vc> -Password '<pw>' -MetaDir .\export-A -Cluster cl01 -Datastore ds01

# Step 4 (new vCenter)
.\Register-VmxFromDatastore.ps1 -Server <new-vc> -User administrator@vsphere.local -Password '<pw>' -Cluster cl01 -UnregisteredCsv .\export-A\unregistered.csv

# Step 5 (new vCenter)
.\Import-VcMeta.ps1 -Server <new-vc> -User administrator@vsphere.local -Password '<pw>' -InDir .\export-A -DatacenterMap '<oldDC>=<newDC>' -Include CustomAttributes,Notes,Tags
```

Add `-DryRun` to any step to see what it would do without writing anything.
Inside PowerShell / ISE the commands can be run as shown; from `cmd.exe` use
`powershell -Command "& .\Script.ps1 ..."` so that comma-separated lists such as
`-Include CustomAttributes,Tags` are passed as arrays.

## Notes

- `-Datastore ds01` in step 3 selects the VMs whose **vmx** is on ds01 (the VM's
  home datastore) and that are listed in the step 0 export. Powered-on VMs are
  skipped unless `-ShutdownFirst`. VMs that span datastores are reported as
  `WarnDiskElsewhere` / `WarnDiskOnDatastore` (warning only).
- `unregistered.csv` is the exact input of step 4 and can be edited in Excel
  first (e.g. change `FolderPath`).
- `Copy-VcMeta.ps1`, `Copy-VcCustomAttributes.ps1`, `Copy-VcFolders.ps1`: one-shot
  A -> B variants (export + import in one run) for when both vCenters are reachable.
- Full documentation (Chinese) and the test report are in the repository root and `docs/`.
