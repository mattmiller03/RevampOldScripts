<#
.SYNOPSIS
    One-time setup script to store vCenter credentials as DPAPI-encrypted XML files.

.DESCRIPTION
    Reads the vCenter list from config/vcenters.json and prompts for credentials for each
    vCenter server. Credentials are stored as encrypted .cred.xml files using PowerShell's
    Export-Clixml, which encrypts via Windows DPAPI (tied to the current user and machine).

    This script must be run interactively (it prompts for credentials) and must be run
    as the same user account that will execute the inventory scripts (e.g., the Task
    Scheduler service account).

.PARAMETER ConfigFile
    Path to the JSON configuration file containing the vCenter list.
    Defaults to config/vcenters.json relative to this script's directory.

.EXAMPLE
    .\Initialize-VCenterCredentials.ps1
    Prompts for credentials for each vCenter defined in config/vcenters.json.

.EXAMPLE
    .\Initialize-VCenterCredentials.ps1 -ConfigFile "C:\Config\vcenters.json"
    Uses a custom config file path.
#>

[CmdletBinding()]
param(
    [Parameter()]
    [string]$ConfigFile = (Join-Path $PSScriptRoot 'config\vcenters.json')
)

$ErrorActionPreference = 'Stop'

# --- Validate config file exists ---
if (-not (Test-Path -Path $ConfigFile)) {
    Write-Error "Configuration file not found: $ConfigFile"
    return
}

$config = Get-Content -Path $ConfigFile -Raw | ConvertFrom-Json
$credDir = Join-Path $PSScriptRoot $config.CredentialDir

# --- Ensure credential directory exists ---
if (-not (Test-Path -Path $credDir)) {
    Write-Host "Creating credential directory: $credDir" -ForegroundColor Cyan
    New-Item -Path $credDir -ItemType Directory -Force | Out-Null
}

# --- Prompt for and store credentials for each vCenter ---
foreach ($vc in $config.VCenters) {
    $credPath = Join-Path $credDir $vc.CredentialFile
    $vcAlias = if ($vc.Alias) { $vc.Alias } else { $vc.Name }
    $vcEnv = if ($vc.TagEnvironment) { " [$($vc.TagEnvironment)]" } else { '' }

    Write-Host "`nEnter credentials for vCenter: $vcAlias$vcEnv ($($vc.Name))" -ForegroundColor Cyan
    Write-Host "  Credential file: $credPath" -ForegroundColor Gray

    $credential = Get-Credential -Message "Credentials for $vcAlias$vcEnv ($($vc.Name))"

    if ($null -eq $credential) {
        Write-Warning "Skipped $vcAlias ($($vc.Name)) - no credential provided."
        continue
    }

    $credential | Export-Clixml -Path $credPath -Force
    Write-Host "  Credential stored for $vcAlias$vcEnv." -ForegroundColor Green
}

Write-Host "`nSetup complete. Stored credentials for $($config.VCenters.Count) vCenter(s) in '$credDir'." -ForegroundColor Green
Write-Host "Credential files are encrypted via Windows DPAPI for the current user on this machine." -ForegroundColor Gray
Write-Host "You can now run Get-AllHostInventory.ps1 and Get-AllVMInventory.ps1 unattended." -ForegroundColor Gray
