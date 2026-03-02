# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Purpose

This repo contains VMware PowerCLI PowerShell scripts for collecting vSphere inventory data. Legacy scripts are kept for reference; revamped versions are the active codebase.

## Repository Structure

```
config/vcenters.json              # vCenter list, credential paths, required tag categories
config/credentials/*.cred.xml     # DPAPI-encrypted credential files (per-user, per-machine)
Initialize-VCenterCredentials.ps1 # One-time setup: prompts for credentials and saves .cred.xml files
Get-AllHostInventory.ps1          # Revamped: ESX host inventory collection (per-vCenter .xlsm)
Get-AllVMInventory.ps1            # Revamped: VM inventory collection (single multi-tab .xlsm)
Get_All_Host_Inventory.ps1        # Legacy (reference only)
Get_All_VM_Inventory.ps1          # Legacy (reference only)
```

## Setup

1. Install modules: `Install-Module VMware.PowerCLI, ImportExcel`
2. Edit `config/vcenters.json` with your vCenter servers, credential filenames, and required tag categories
3. Run `.\Initialize-VCenterCredentials.ps1` interactively as the service account — prompts for credentials per vCenter and saves DPAPI-encrypted .cred.xml files
4. Schedule `Get-AllHostInventory.ps1` and/or `Get-AllVMInventory.ps1` via Task Scheduler

## Architecture

### Credential Storage
Credentials are stored as DPAPI-encrypted XML files via PowerShell's `Export-Clixml` / `Import-Clixml`. No extra modules required — encryption is tied to the current Windows user account and machine. The `config/credentials/` directory holds one `.cred.xml` file per vCenter.

### Host Inventory (Get-AllHostInventory.ps1)
- Produces one `.xlsm` workbook per vCenter with Search tab and HostInventory data tab
- Same collect/export pattern as before

### VM Inventory (Get-AllVMInventory.ps1)
- Produces a **single** `VMInventory_All.xlsm` workbook with multiple tabs:
  - **Search** — VBA-powered search UI with Add Entry functionality (searches All_VMs table)
  - **All_VMs** — Combined VM inventory from all vCenters
  - **MissingTags** — VMs missing any required tag category (configured in JSON)
  - **VM_BIOS** — VMs using BIOS firmware (not EFI)
  - **VMs_Powered_Off** — VMs in PoweredOff state
  - **\<vCenter name\>** — One tab per vCenter with that vCenter's VMs

### Tag Configuration
Tag categories and column counts are driven by `RequiredTags` in `config/vcenters.json`:
```json
"RequiredTags": [
    { "Category": "Application", "Columns": 2 },
    { "Category": "VlanID", "Columns": 4 }
]
```
Tag columns are built dynamically on each VM object (e.g., `Application_Tag1`, `Application_Tag2`). A VM appears on the MissingTags tab if **any** required tag category has all its columns blank.

### Shared Patterns
Both inventory scripts share:
- `[CmdletBinding()]` with `param()` block — all paths configurable, defaults relative to `$PSScriptRoot`
- Config loaded from JSON (`config/vcenters.json`) via `ConvertFrom-Json`
- Credentials loaded via `Import-Clixml` from `config/credentials/`
- `Backup-PreviousReport` function archives prior workbook before each run
- `try/catch/finally` per vCenter: connect, collect inventory, disconnect
- Connection object captured from `Connect-VIServer` and disconnected specifically (no wildcard `*`)
- Timestamped transcripts written to `Output/Transcripts/`
- Summary at end: success/fail counts, duration

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
