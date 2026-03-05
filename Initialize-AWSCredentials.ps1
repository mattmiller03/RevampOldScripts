<#
.SYNOPSIS
    One-time setup script to store AWS credentials as DPAPI-encrypted XML files.

.DESCRIPTION
    Reads the AWS region/credential configuration from config/aws.json and prompts for
    AWS access key credentials. Credentials are stored as an encrypted .cred.xml file
    using PowerShell's Export-Clixml, which encrypts via Windows DPAPI (tied to the
    current user and machine).

    The PSCredential stores:
      - Username = AWS Access Key ID
      - Password = AWS Secret Access Key

    This script must be run interactively (it prompts for credentials) and must be run
    as the same user account that will execute the inventory scripts (e.g., the Task
    Scheduler service account).

.PARAMETER ConfigFile
    Path to the JSON configuration file containing the AWS region list.
    Defaults to config/aws.json relative to this script's directory.

.PARAMETER TestConnection
    If specified, validates the credential by listing EC2 instances in the first
    configured region, then clears the session credential.

.EXAMPLE
    .\Initialize-AWSCredentials.ps1
    Prompts for AWS access key credentials and saves the DPAPI-encrypted .cred.xml file.

.EXAMPLE
    .\Initialize-AWSCredentials.ps1 -TestConnection
    Prompts for credentials and validates by connecting to the first configured region.
#>

[CmdletBinding()]
param(
    [Parameter()]
    [string]$ConfigFile = (Join-Path $PSScriptRoot 'config\aws.json'),

    [Parameter()]
    [switch]$TestConnection
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

# --- Prompt for and store credential ---
$credPath = Join-Path $credDir $config.CredentialFile
$regionList = ($config.Regions | ForEach-Object { $_.Alias }) -join ', '

Write-Host "`nEnter AWS credentials for regions: $regionList" -ForegroundColor Cyan
Write-Host "  Username = AWS Access Key ID" -ForegroundColor Gray
Write-Host "  Password = AWS Secret Access Key" -ForegroundColor Gray
Write-Host "  Credential file: $credPath" -ForegroundColor Gray

$credential = Get-Credential -Message "AWS Credentials (Username = Access Key ID, Password = Secret Access Key)"

if ($null -eq $credential) {
    Write-Warning "No credential provided. Exiting."
    return
}

$credential | Export-Clixml -Path $credPath -Force
Write-Host "  Credential stored." -ForegroundColor Green

# Optionally test the connection
if ($TestConnection) {
    $testRegion = $config.Regions[0].Name
    Write-Host "  Testing connection to $testRegion..." -ForegroundColor Gray
    try {
        Set-AWSCredential -AccessKey $credential.UserName `
            -SecretKey $credential.GetNetworkCredential().Password
        Set-DefaultAWSRegion -Region $testRegion
        $null = Get-EC2Instance -ErrorAction Stop
        Write-Host "  Connection successful." -ForegroundColor Green
        Clear-AWSCredential
        Clear-DefaultAWSRegion
    }
    catch {
        Write-Warning "  Connection test failed: $_"
    }
}

Write-Host "`nSetup complete. Credential file stored in '$credDir'." -ForegroundColor Green
Write-Host "Credential file is encrypted via Windows DPAPI for the current user on this machine." -ForegroundColor Gray
Write-Host "You can now run Get-AllAWSInventory.ps1 unattended." -ForegroundColor Gray
