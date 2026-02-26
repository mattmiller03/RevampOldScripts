# RevampOldScripts

Taking old scripts and updating them with readable coding standards.

## Setup

1. Install required modules:
   ```powershell
   Install-Module VMware.PowerCLI, Microsoft.PowerShell.SecretManagement, Microsoft.PowerShell.SecretStore
   ```
2. Edit `config/vcenters.json` with your vCenter servers
3. Run `.\Initialize-VCenterSecrets.ps1` interactively to store encrypted credentials
4. Schedule `Get-AllHostInventory.ps1` and/or `Get-AllVMInventory.ps1` via Task Scheduler

## Project TODO

### Completed
- [x] Review legacy scripts for bugs and issues
- [x] Create JSON config structure (`config/vcenters.json`)
- [x] Create credential setup script (`Initialize-VCenterSecrets.ps1`) using SecretManagement
- [x] Rewrite host inventory script (`Get-AllHostInventory.ps1`)
- [x] Rewrite VM inventory script (`Get-AllVMInventory.ps1`)
- [x] Copy script-analyzer and powercli-helper agents to this repo
- [x] Update CLAUDE.md with new architecture
- [x] Set up `.claude/settings.local.json` with PowerShell MCP permissions
- [x] Validate new scripts - PSScriptAnalyzer clean, zero parse errors
- [x] Fill in host inventory collection logic (51 columns matching original report)
- [x] Add `.gitignore` for Output directories and credential files
- [x] Move legacy scripts to `legacy/` folder
- [x] Fill in VM inventory collection logic (68 columns matching original report)

### TODO
- [ ] Test `Initialize-VCenterSecrets.ps1` against a real vault
- [ ] Test host inventory script against a live vCenter
- [ ] Test VM inventory script against a live vCenter

### Notes
- `EBS_Number`, `DLA_Asset`, `Site_Location` columns in the host report are output as empty strings - these appear to be custom fields that need to be mapped to your environment's custom attributes or tags
- The VM report property list came from an xlsx file (`VMpropertylist.csv`) - 68 columns including tags, disk sizes, and vNICs

## File Structure

```
config/vcenters.json             # vCenter list and secret names
Initialize-VCenterSecrets.ps1    # One-time credential setup (interactive)
Get-AllHostInventory.ps1         # ESX host inventory (revamped, 51 columns)
Get-AllVMInventory.ps1           # VM inventory (revamped, 68 columns)
.gitignore                       # Excludes Output/, *.xlsx, *.cred.xml
legacy/                          # Original scripts kept for reference
  Get_All_Host_Inventory.ps1
  Get_All_VM_Inventory.ps1
.claude/agents/                  # Claude Code custom agents
  script-analyzer.md             # Code review and analysis agent
  powercli-helper.md             # PowerCLI command generation agent
```
