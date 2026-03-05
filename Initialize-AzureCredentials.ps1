<#
.SYNOPSIS
    One-time setup script to store Azure service principal credentials as DPAPI-encrypted XML files.

.DESCRIPTION
    Reads the Azure tenant/subscription list from config/azure.json and prompts for service
    principal credentials for each unique credential file. Credentials are stored as encrypted .cred.xml
    files using PowerShell's Export-Clixml, which encrypts via Windows DPAPI (tied to the current
    user and machine).

    The PSCredential stores:
      - Username = Azure Application (Client) ID
      - Password = Azure Client Secret

    This script must be run interactively (it prompts for credentials) and must be run as the
    same user account that will execute the inventory scripts (e.g., the Task Scheduler service
    account).

.PARAMETER ConfigFile
    Path to the JSON configuration file containing the Azure tenant/subscription list.
    Defaults to config/azure.json relative to this script's directory.

.PARAMETER TestConnection
    If specified, validates each credential by connecting to Azure and then disconnecting.

.EXAMPLE
    .\Initialize-AzureCredentials.ps1
    Prompts for credentials for each unique credential file defined in config/azure.json.

.EXAMPLE
    .\Initialize-AzureCredentials.ps1 -TestConnection
    Prompts for credentials and validates each by connecting to Azure.
#>

[CmdletBinding()]
param(
    [Parameter()]
    [string]$ConfigFile = (Join-Path $PSScriptRoot 'config\azure.json'),

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

# --- Deduplicate credential files (multiple tenants may share the same credential) ---
$credentialMap = [ordered]@{}
foreach ($tenant in $config.Tenants) {
    $credFile = $tenant.CredentialFile
    if (-not $credentialMap.ContainsKey($credFile)) {
        $credentialMap[$credFile] = @{
            TenantIDs = [System.Collections.Generic.List[string]]::new()
            SubAliases = [System.Collections.Generic.List[string]]::new()
        }
    }
    $credentialMap[$credFile].TenantIDs.Add($tenant.TenantID)
    foreach ($sub in $tenant.Subscriptions) {
        $subAlias = if ($sub.Alias) { $sub.Alias } else { $sub.Name }
        $credentialMap[$credFile].SubAliases.Add($subAlias)
    }
}

# --- Prompt for and store credentials for each unique credential file ---
foreach ($credFile in $credentialMap.Keys) {
    $credPath = Join-Path $credDir $credFile
    $subList = $credentialMap[$credFile].SubAliases -join ', '
    $tenantList = $credentialMap[$credFile].TenantIDs -join ', '

    Write-Host "`nEnter Azure Service Principal credentials for: $subList" -ForegroundColor Cyan
    Write-Host "  Tenant(s): $tenantList" -ForegroundColor Gray
    Write-Host "  Username = Application (Client) ID" -ForegroundColor Gray
    Write-Host "  Password = Client Secret" -ForegroundColor Gray
    Write-Host "  Credential file: $credPath" -ForegroundColor Gray

    $credential = Get-Credential -Message "Azure SP for $subList (Username = AppID, Password = ClientSecret)"

    if ($null -eq $credential) {
        Write-Warning "Skipped $subList - no credential provided."
        continue
    }

    $credential | Export-Clixml -Path $credPath -Force
    Write-Host "  Credential stored for $subList." -ForegroundColor Green

    # Optionally test the connection (uses the first tenant for this credential)
    if ($TestConnection) {
        $testTenantId = $credentialMap[$credFile].TenantIDs[0]
        Write-Host "  Testing connection to $($config.Environment) (Tenant: $testTenantId)..." -ForegroundColor Gray
        try {
            Connect-AzAccount -Environment $config.Environment -ServicePrincipal `
                -TenantId $testTenantId -Credential $credential -ErrorAction Stop | Out-Null
            Write-Host "  Connection successful." -ForegroundColor Green
            Disconnect-AzAccount -ErrorAction SilentlyContinue | Out-Null
        }
        catch {
            Write-Warning "  Connection test failed: $_"
        }
    }
}

$uniqueCount = $credentialMap.Keys.Count
$subCount = ($config.Tenants.Subscriptions | Measure-Object).Count
$tenantCount = $config.Tenants.Count
Write-Host "`nSetup complete. Stored $uniqueCount credential file(s) for $tenantCount tenant(s) / $subCount subscription(s) in '$credDir'." -ForegroundColor Green
Write-Host "Credential files are encrypted via Windows DPAPI for the current user on this machine." -ForegroundColor Gray
Write-Host "You can now run Get-AllAzureInventory.ps1 unattended." -ForegroundColor Gray
