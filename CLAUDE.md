# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Purpose

This repo contains VMware PowerCLI PowerShell scripts for collecting vSphere inventory data. Legacy scripts are kept for reference; revamped versions are the active codebase.

## Repository Structure

```
config/vcenters.json           # vCenter list + SecretManagement vault/secret names
Initialize-VCenterSecrets.ps1  # One-time setup: creates vault and stores credentials
Get-AllHostInventory.ps1       # Revamped: ESX host inventory collection
Get-AllVMInventory.ps1         # Revamped: VM inventory collection
Get_All_Host_Inventory.ps1     # Legacy (reference only)
Get_All_VM_Inventory.ps1       # Legacy (reference only)
```

## Setup

1. Install modules: `Install-Module VMware.PowerCLI, Microsoft.PowerShell.SecretManagement, Microsoft.PowerShell.SecretStore`
2. Edit `config/vcenters.json` with your vCenter servers and desired secret names
3. Run `.\Initialize-VCenterSecrets.ps1` interactively as the service account — prompts for credentials per vCenter and stores them encrypted in a SecretStore vault
4. Schedule `Get-AllHostInventory.ps1` and/or `Get-AllVMInventory.ps1` via Task Scheduler

## Architecture

Both inventory scripts share the same pattern:
- `[CmdletBinding()]` with `param()` block — all paths configurable, defaults relative to `$PSScriptRoot`
- Config loaded from JSON (`config/vcenters.json`) via `ConvertFrom-Json`
- Credentials retrieved from SecretManagement vault as `[PSCredential]` objects
- `Backup-PreviousReport` function archives prior CSV before each run (distinct source/archive paths)
- `try/catch/finally` per vCenter: connect, collect inventory, save CSV, disconnect
- Connection object captured from `Connect-VIServer` and disconnected specifically (no wildcard `*`)
- Timestamped transcripts written to `Output/Transcripts/`
- Summary at end: success/fail counts, duration

The inventory collection blocks are currently **stubbed with TODO placeholders** — replace with your actual `Get-VMHost` / `Get-VM` pipeline logic.

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
