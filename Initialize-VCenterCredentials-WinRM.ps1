<#
.SYNOPSIS
    Re-encrypts vCenter credentials via WinRM so they work under Aria Orchestrator.

.DESCRIPTION
    DPAPI-encrypted credential files created during an interactive (RDP) session cannot
    be decrypted under a WinRM network logon, even with the same service account.
    This script collects credentials locally (GUI prompts), then creates the encrypted
    .cred.xml files inside a WinRM session so the DPAPI keys match the logon type that
    Aria Orchestrator's PowerShell plugin uses.

.PARAMETER PowerShellHost
    The hostname or FQDN of the Aria PowerShell host to connect to via WinRM.

.PARAMETER ServiceAccount
    The service account username (domain\user) that Aria uses to connect to the PS host.

.PARAMETER ScriptRoot
    The root directory on the remote host where the inventory scripts and config live.

.PARAMETER ConfigFile
    Path to the JSON config file relative to ScriptRoot. Defaults to config\vcenters.json.

.EXAMPLE
    .\Initialize-VCenterCredentials-WinRM.ps1 -PowerShellHost "pshost.domain.com" -ServiceAccount "DOMAIN\svc_vrapsh"

.EXAMPLE
    .\Initialize-VCenterCredentials-WinRM.ps1 -PowerShellHost "pshost.domain.com" -ServiceAccount "DOMAIN\svc_vrapsh" -ScriptRoot "D:\Scripts\Inventory"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$PowerShellHost,

    [Parameter(Mandatory)]
    [string]$ServiceAccount,

    [Parameter()]
    [string]$ScriptRoot = 'C:\DLA-Failsafe\vRA\MattM\Workflow\Inventory',

    [Parameter()]
    [string]$ConfigFile = 'config\vcenters.json'
)

$ErrorActionPreference = 'Stop'

# --- Read the config locally to know which vCenters need credentials ---
# We read a local copy to display prompts; the remote session reads its own copy
$localConfigPath = Join-Path $PSScriptRoot $ConfigFile
if (-not (Test-Path -Path $localConfigPath)) {
    Write-Warning "Local config not found at '$localConfigPath'. You will be prompted for vCenter names manually."
    $vcNames = @()
    while ($true) {
        $name = Read-Host "Enter a vCenter FQDN (or blank to finish)"
        if ([string]::IsNullOrWhiteSpace($name)) { break }
        $vcNames += $name
    }
    if ($vcNames.Count -eq 0) {
        Write-Error "No vCenters specified. Exiting."
        return
    }
}
else {
    $localConfig = Get-Content -Path $localConfigPath -Raw | ConvertFrom-Json
    $vcNames = $localConfig.VCenters | ForEach-Object { $_.Name }
    Write-Host "Found $($vcNames.Count) vCenter(s) in config:" -ForegroundColor Cyan
    foreach ($name in $vcNames) {
        Write-Host "  - $name" -ForegroundColor Gray
    }
}

# --- Collect credentials locally (GUI prompts work here) ---
Write-Host "`nCollecting credentials locally..." -ForegroundColor Cyan
$vcCreds = @{}
foreach ($vcName in $vcNames) {
    Write-Host "`nEnter credentials for: $vcName" -ForegroundColor Cyan
    $cred = Get-Credential -Message "Credentials for vCenter: $vcName"
    if ($null -eq $cred) {
        Write-Warning "Skipped $vcName - no credential provided."
        continue
    }
    $vcCreds[$vcName] = $cred
}

if ($vcCreds.Count -eq 0) {
    Write-Error "No credentials collected. Exiting."
    return
}

# --- Connect to the remote PS host via WinRM ---
Write-Host "`nConnecting to $PowerShellHost as $ServiceAccount via WinRM..." -ForegroundColor Cyan
$session = New-PSSession -ComputerName $PowerShellHost -Credential (Get-Credential -UserName $ServiceAccount -Message "Service account credentials for WinRM session")

try {
    # --- Create credential files inside the WinRM session ---
    $result = Invoke-Command -Session $session -ScriptBlock {
        param($creds, $remoteScriptRoot, $remoteConfigFile)

        $configPath = Join-Path $remoteScriptRoot $remoteConfigFile
        if (-not (Test-Path -Path $configPath)) {
            throw "Remote config file not found: $configPath"
        }

        $config = Get-Content -Path $configPath -Raw | ConvertFrom-Json
        $credDir = Join-Path $remoteScriptRoot $config.CredentialDir

        if (-not (Test-Path -Path $credDir)) {
            New-Item -Path $credDir -ItemType Directory -Force | Out-Null
            Write-Host "Created credential directory: $credDir"
        }

        $stored = 0
        $skipped = 0
        foreach ($vc in $config.VCenters) {
            $credPath = Join-Path $credDir $vc.CredentialFile
            if ($creds.ContainsKey($vc.Name)) {
                $creds[$vc.Name] | Export-Clixml -Path $credPath -Force
                Write-Host "Stored credential for $($vc.Name) -> $credPath"
                $stored++
            }
            else {
                Write-Warning "No credential provided for $($vc.Name) - skipped"
                $skipped++
            }
        }

        [PSCustomObject]@{ Stored = $stored; Skipped = $skipped }
    } -ArgumentList $vcCreds, $ScriptRoot, $ConfigFile

    Write-Host "`nDone. Stored: $($result.Stored), Skipped: $($result.Skipped)" -ForegroundColor Green
    Write-Host "Credential files are now DPAPI-encrypted under the WinRM network logon context." -ForegroundColor Gray
    Write-Host "They will work when Aria Orchestrator runs scripts via the PowerShell plugin." -ForegroundColor Gray
}
finally {
    Remove-PSSession $session
}
