<#
.SYNOPSIS
    Sanitizes legacy PowerShell scripts by replacing sensitive content with generic placeholders.

.DESCRIPTION
    Reads legacy scripts from an input directory, applies an ordered set of regex-based
    replacement rules to remove FQDNs, hostnames, usernames, file paths, credential logging,
    and infrastructure references. Writes sanitized copies to a separate output directory.

    Original files are never modified. The operation is idempotent — running it multiple
    times produces identical output.

.PARAMETER InputDir
    Directory containing the legacy scripts to sanitize. Defaults to legacy/ relative to
    this script's directory.

.PARAMETER OutputDir
    Directory where sanitized copies are written. Created if it does not exist.
    Defaults to legacy/scrubbed/ relative to this script's directory.

.EXAMPLE
    .\Scrub-LegacyScripts.ps1
    Sanitizes all .ps1 files in legacy/ and writes output to legacy/scrubbed/.

.EXAMPLE
    .\Scrub-LegacyScripts.ps1 -OutputDir "C:\CleanScripts" -Verbose
    Sanitizes with a custom output path and verbose logging.
#>

[CmdletBinding()]
param(
    [Parameter()]
    [string]$InputDir = (Join-Path $PSScriptRoot 'legacy'),

    [Parameter()]
    [string]$OutputDir = (Join-Path $PSScriptRoot 'legacy\scrubbed')
)

$ErrorActionPreference = 'Stop'

#region Functions

function Invoke-ContentScrub {
    <#
    .SYNOPSIS
        Applies ordered regex replacement rules to a string and returns match counts.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Content,

        [Parameter(Mandatory)]
        [PSCustomObject[]]$Rules
    )

    $regexOpts = [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor
                 [System.Text.RegularExpressions.RegexOptions]::Multiline

    $replacementCounts = [ordered]@{}

    foreach ($rule in $Rules) {
        $matchCount = ([regex]::Matches($Content, $rule.Pattern, $regexOpts)).Count
        $replacementCounts[$rule.Name] = $matchCount

        if ($matchCount -gt 0) {
            $Content = [regex]::Replace($Content, $rule.Pattern, $rule.Replacement, $regexOpts)
            Write-Verbose "  $($rule.Name): $matchCount replacement(s)"
        }
    }

    [PSCustomObject]@{
        Content           = $Content
        ReplacementCounts = $replacementCounts
    }
}

#endregion Functions

#region Validation

if (-not (Test-Path -Path $InputDir)) {
    Write-Error "Input directory not found: $InputDir"
    return
}

if (-not (Test-Path -Path $OutputDir)) {
    New-Item -Path $OutputDir -ItemType Directory -Force | Out-Null
    Write-Verbose "Created output directory: $OutputDir"
}

#endregion Validation

#region Replacement Rules

# Order matters: line-level removals first, then compound strings, then specific tokens,
# then numbered patterns before bare, then paths from most-specific to least-specific.
$Rules = @(
    # --- Phase 1: Remove credential logging lines entirely ---
    [PSCustomObject]@{
        Name        = 'CredentialLogging'
        Pattern     = '^\s*write-host\s+"Credentials returned\.\.\."\s+\$creds\.Host\s+\$creds\.User\s+\$creds\.Password\s*$'
        Replacement = '        # REMOVED: Line logged credentials to console'
    }

    # --- Phase 2: Compound comment/string replacements ---
    [PSCustomObject]@{
        Name        = 'HawaiiHostname'
        Pattern     = 'HAWAII vCenter hnl1s-ph1005v not working'
        Replacement = 'remote site vCenter not reachable'
    }
    [PSCustomObject]@{
        Name        = 'JumpServerComment'
        Pattern     = 'copied from vs346 and modified to run on the new Windows 2016 jump server'
        Replacement = 'copied from previous server and modified to run on the new management server'
    }
    [PSCustomObject]@{
        Name        = 'DirMAccount'
        Pattern     = 'DIR M account'
        Replacement = 'service account'
    }
    [PSCustomObject]@{
        Name        = 'DomainAccountComment'
        Pattern     = 'run using a domain account'
        Replacement = 'run using a service account'
    }
    [PSCustomObject]@{
        Name        = 'AuthorInitials'
        Pattern     = '#JN \d{1,2}/\d{1,2}/\d{4}'
        Replacement = '# Author MM/DD/YYYY'
    }

    # --- Phase 3: Credential and account replacements ---
    [PSCustomObject]@{
        Name        = 'SsoAdminDomain'
        Pattern     = '"administrator@ssodomain"'
        Replacement = '"administrator@vsphere.local"'
    }
    [PSCustomObject]@{
        Name        = 'SsoAdminAdmin'
        Pattern     = '"administrator@ssoadmin"'
        Replacement = '"administrator@vsphere.local"'
    }
    [PSCustomObject]@{
        Name        = 'DomainBackslashUser'
        Pattern     = 'domain\\user'
        Replacement = 'CORP\svc_vcenter'
    }

    # --- Phase 4: File path replacements ---
    [PSCustomObject]@{
        Name        = 'CredStorePath'
        Pattern     = 'C:\\Users\\AppData\\Roaming\\VMware\\credstore\\vicredentials\.xml'
        Replacement = 'C:\Users\svc_vcenter\AppData\Roaming\VMware\credstore\vicredentials.xml'
    }

    # --- Phase 5: vCenter name replacements (numbered before bare) ---
    [PSCustomObject]@{
        Name        = 'NameOfVcenterNumbered'
        Pattern     = 'nameofvcenter(\d+)'
        Replacement = 'vcenter0${1}.corp.example.com'
    }
    [PSCustomObject]@{
        Name        = 'NameOfVcenterBare'
        Pattern     = 'nameofvcenter'
        Replacement = 'vcenter01.corp.example.com'
    }
    [PSCustomObject]@{
        Name        = 'VcenterName3'
        Pattern     = 'vcentername3'
        Replacement = 'vcenter03.corp.example.com'
    }
    [PSCustomObject]@{
        Name        = 'VcenterName2'
        Pattern     = 'vcentername2'
        Replacement = 'vcenter02.corp.example.com'
    }
    [PSCustomObject]@{
        Name        = 'VcenterName1'
        Pattern     = 'vcentername1'
        Replacement = 'vcenter01.corp.example.com'
    }

    # --- Phase 6: Drive letter paths (most specific to least) ---
    [PSCustomObject]@{
        Name        = 'DrivePathEScriptOutput'
        Pattern     = '[Ee]:\\Script_Output'
        Replacement = 'D:\Script_Output'
    }
    [PSCustomObject]@{
        Name        = 'DrivePathEScriptsProduction'
        Pattern     = '[Ee]:\\Scripts\\PRODUCTION'
        Replacement = 'D:\Scripts\PRODUCTION'
    }
    [PSCustomObject]@{
        Name        = 'DrivePathEScripts'
        Pattern     = '[Ee]:\\Scripts'
        Replacement = 'D:\Scripts'
    }
    [PSCustomObject]@{
        Name        = 'DrivePathEGeneric'
        Pattern     = '[Ee]:\\'
        Replacement = 'D:\'
    }
    [PSCustomObject]@{
        Name        = 'DrivePathENoSlash'
        Pattern     = '(?<=[" ])E:'
        Replacement = 'D:'
    }
)

#endregion Replacement Rules

#region Main

$scriptFiles = Get-ChildItem -Path $InputDir -Filter '*.ps1' -File
if ($scriptFiles.Count -eq 0) {
    Write-Warning "No .ps1 files found in: $InputDir"
    return
}

Write-Host "Scrub-LegacyScripts - Sanitizing legacy scripts for safe sharing" -ForegroundColor Cyan
Write-Host ("=" * 65) -ForegroundColor Cyan
Write-Host "Input:  $InputDir"
Write-Host "Output: $OutputDir"
Write-Host ""

$totalFiles = 0

foreach ($file in $scriptFiles) {
    Write-Host "Processing: $($file.Name)" -ForegroundColor Yellow
    $content = Get-Content -Path $file.FullName -Raw -Encoding UTF8

    $result = Invoke-ContentScrub -Content $content -Rules $Rules

    $outputPath = Join-Path $OutputDir $file.Name
    Set-Content -Path $outputPath -Value $result.Content -NoNewline -Encoding UTF8

    # Report per-rule counts (only rules that matched)
    $fileTotal = 0
    $activeRules = 0
    foreach ($key in $result.ReplacementCounts.Keys) {
        $count = $result.ReplacementCounts[$key]
        if ($count -gt 0) {
            $label = "{0,-30}" -f $key
            Write-Host "  $label $count replacement(s)" -ForegroundColor Gray
            $fileTotal += $count
            $activeRules++
        }
    }
    Write-Host "  Total: $fileTotal replacements across $activeRules rules" -ForegroundColor Green
    Write-Host ""
    $totalFiles++
}

Write-Host ("=" * 65) -ForegroundColor Cyan
Write-Host "Complete. $totalFiles file(s) sanitized to $OutputDir" -ForegroundColor Green

#endregion Main
