# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Purpose

This repo contains PowerShell scripts for collecting infrastructure inventory data from VMware vSphere, Azure, and AWS cloud environments. Legacy scripts are kept for reference; revamped versions are the active codebase.

## Repository Structure

```
config/vcenters.json              # vCenter list, credential paths, required tag categories
config/azure.json                 # Azure subscription list, tenant, required tags
config/aws.json                   # AWS region list, credential path
config/credentials/*.cred.xml     # DPAPI-encrypted credential files (per-user, per-machine)
Initialize-VCenterCredentials.ps1 # One-time setup: prompts for vCenter credentials
Initialize-AzureCredentials.ps1   # One-time setup: prompts for Azure SP credentials
Initialize-AWSCredentials.ps1     # One-time setup: prompts for AWS access key credentials
Get-AllHostInventory.ps1          # Revamped: ESX host inventory collection (per-vCenter .xlsm)
Get-AllVMInventory.ps1            # Revamped: VM inventory collection (single multi-tab .xlsm)
Get-AllAzureInventory.ps1         # Revamped: Azure VM inventory collection (single multi-tab .xlsm)
Get-AllAWSInventory.ps1           # Revamped: AWS EC2 inventory collection (single multi-tab .xlsm)
Get_All_Host_Inventory.ps1        # Legacy (reference only)
Get_All_VM_Inventory.ps1          # Legacy (reference only)
```

## Setup

1. Install modules: `Install-Module VMware.PowerCLI, ImportExcel`
2. Edit `config/vcenters.json` with your vCenter servers, credential filenames, and required tag categories
3. Run `.\Initialize-VCenterCredentials.ps1` interactively as the service account — prompts for credentials per vCenter and saves DPAPI-encrypted .cred.xml files
4. Schedule `Get-AllHostInventory.ps1` and/or `Get-AllVMInventory.ps1` via Task Scheduler
5. For Azure: `Install-Module Az` then edit `config/azure.json` with your subscriptions, tenant ID, and credential filenames
6. Run `.\Initialize-AzureCredentials.ps1` interactively — prompts for service principal credentials (AppID + ClientSecret)
7. Schedule `Get-AllAzureInventory.ps1` via Task Scheduler
8. For AWS: `Install-Module AWS.Tools.Common, AWS.Tools.EC2` then edit `config/aws.json` with your regions and credential filename
9. Run `.\Initialize-AWSCredentials.ps1` interactively — prompts for AWS access key credentials (AccessKeyID + SecretAccessKey)
10. Schedule `Get-AllAWSInventory.ps1` via Task Scheduler

## Architecture

### Credential Storage
Credentials are stored as DPAPI-encrypted XML files via PowerShell's `Export-Clixml` / `Import-Clixml`. No extra modules required — encryption is tied to the current Windows user account and machine. The `config/credentials/` directory holds one `.cred.xml` file per vCenter.

### Host Inventory (Get-AllHostInventory.ps1)
- Produces one `.xlsm` workbook per vCenter with multiple tabs:
  - **Search** — VBA-powered search UI
  - **All_Hosts** — Combined host inventory from all vCenters
  - **NOT_SecureBoot** — Hosts without Secure Boot enabled
  - **Not_Patched** — Hosts not on the target ESXi version/build (requires `TargetESXiVersion`/`TargetESXiBuild` in config)
  - **Not_Connected** — Hosts not in Connected state
  - **Not_On_ESX_\<N\>** — Hosts not on the target major ESXi version
  - **\<vCenter name\>** — One tab per vCenter with that vCenter's hosts

### VM Inventory (Get-AllVMInventory.ps1)
- Produces a **single** `VMInventory_All.xlsm` workbook with multiple tabs:
  - **Search** — VBA-powered search UI (searches All_VMs table)
  - **All_VMs** — Combined VM inventory from all vCenters
  - **MissingTags** — VMs missing any required tag category (configured in JSON)
  - **VM_BIOS** — VMs using BIOS firmware (not EFI)
  - **VMs_Powered_Off** — VMs in PoweredOff state
  - **CPU_HotAdd_FALSE** — VMs with CPU Hot Add disabled
  - **Memory_HotAdd_FALSE** — VMs with Memory Hot Add disabled
  - **VMToolsVersion** — VMs not on the latest VMware Tools version
  - **FloppyDrives** — VMs with a floppy drive attached
  - **\<vCenter name\>** — One tab per vCenter with that vCenter's VMs

### ESXi Target Version (Host Inventory)
The host inventory script uses two optional config fields to populate the **Not_Patched** tab:
- `TargetESXiVersion` — expected ESXi version string (e.g., `"8.0.3"`)
- `TargetESXiBuild` — expected ESXi build number (e.g., `"24322831"`)

If either field is missing, the Not_Patched tab is skipped. The major version number (first segment) is also used for the **Not_On_ESX_\<N\>** tab.

### Azure VM Inventory (Get-AllAzureInventory.ps1)
- Targets Azure US Government subscriptions via service principal authentication
- Produces a **single** `AzureInventory_All.xlsm` workbook with multiple tabs:
  - **Search** — VBA-powered search UI (searches All_VMs table)
  - **All_VMs** — Combined Azure VM inventory from all subscriptions
  - **Not_Running** — VMs not in 'running' state
  - **NO_App-name** — VMs missing the App-name tag
  - **NO_BootDiag** — VMs without boot diagnostics configured
  - **IL4_VMs** — VMs with Impact-level tag = IL4
  - **IL5_VMs** — VMs with Impact-level tag = IL5
  - **\<Subscription\>** — One tab per subscription

### AWS EC2 Inventory (Get-AllAWSInventory.ps1)
- Targets AWS GovCloud regions via access key credentials
- Produces a **single** `AWSInventory_All.xlsm` workbook with multiple tabs:
  - **Search** — VBA-powered search UI (searches All_VMs table)
  - **All_VMs** — Combined EC2 inventory from all regions
  - **Not_Running** — Instances not in 'running' state
  - **\<Region\>** — One tab per region (e.g., EAST, WEST)

### Azure Authentication
Azure uses service principal credentials (AppID + ClientSecret) stored as DPAPI-encrypted PSCredential files. The PSCredential Username = ApplicationID, Password = ClientSecret. TenantID and Environment are in `config/azure.json`. A single `Connect-AzAccount` session is shared across all subscriptions, with `Set-AzContext` switching per subscription.

### AWS Authentication
AWS uses access key credentials (AccessKeyID + SecretAccessKey) stored as DPAPI-encrypted PSCredential files. The PSCredential Username = Access Key ID, Password = Secret Access Key. Regions are configured in `config/aws.json`. Credentials are loaded once via `Set-AWSCredential` and `Set-DefaultAWSRegion` switches per region.

### Tag Configuration
vSphere tag categories follow the naming pattern `{TagPrefix}-{TagEnvironment}-{Category}` (e.g., `vCenter-Prod-App-Name`). The JSON config defines:
- `TagPrefix` — global prefix (e.g., `"vCenter"`)
- `RequiredTags` — each with `Category` (vSphere suffix), `DisplayName` (spreadsheet column), and `Columns` count
- `TagEnvironment` — per-vCenter environment identifier (e.g., `"Prod"`, `"Dev"`, `"OT"`)

```json
"TagPrefix": "vCenter",
"RequiredTags": [
    { "Category": "App-Name", "DisplayName": "Application", "Columns": 2 },
    { "Category": "VLAN_ID", "DisplayName": "VlanID", "Columns": 4 }
]
```
At runtime the script looks up `vCenter-Prod-App-Name` in vSphere but writes the column as `Application_Tag1`. A VM appears on the MissingTags tab if **any** required tag category has all its columns blank.

### Shared Patterns
All inventory scripts share:
- `[CmdletBinding()]` with `param()` block — all paths configurable, defaults relative to `$PSScriptRoot`
- Config loaded from JSON via `ConvertFrom-Json`
- Credentials loaded via `Import-Clixml` from `config/credentials/`
- `Backup-PreviousReport` function archives prior workbook before each run
- `try/catch/finally` per target: connect, collect inventory, disconnect
- Timestamped transcripts written to `Output/Transcripts/`
- Summary at end: success/fail counts, duration
- Search tab with VBA macro for filtering All_VMs table
- `[ordered]@{}` cast to `[PSCustomObject]` for consistent column ordering
- `[System.Collections.Generic.List[PSCustomObject]]` for data accumulation

## Conventions

- PowerShell verb-noun naming, PascalCase for functions/parameters
- `$ErrorActionPreference = 'Stop'` at script level; `-ErrorAction Stop` on critical cmdlets
- `Write-Verbose` for debug info, `Write-Warning` for errors, `Write-Host` for user-facing status only
- Never log credentials to any output stream
- `Move-Item`/`Remove-Item` with full cmdlet names and named parameters
- 4-space indentation, no tabs
- `#Requires` directives for module dependencies
- ISO 8601 timestamps (`yyyy-MM-dd HH:mm:ss`)

## Legacy Scripts (reference only)

The `Get_All_*` scripts are the originals kept for reference. They have critical bugs (unparsed CSV, identical source/dest paths, credential logging, zero error handling) and should not be run in production. See git history for the full review.
