#Requires -Modules AWS.Tools.Common, AWS.Tools.EC2, ImportExcel

<#
.SYNOPSIS
    Collects AWS EC2 instance inventory across all configured regions and produces
    a multi-tab Excel workbook.

.DESCRIPTION
    Iterates through each AWS region listed in config/aws.json, authenticates using
    DPAPI-encrypted credentials (created by Initialize-AWSCredentials.ps1), collects
    EC2 instance inventory, and produces a single Excel workbook with multiple tabs:

      - Search       : VBA-powered search UI (searches All_VMs table)
      - All_VMs      : Combined EC2 inventory from all regions
      - Not_Running  : Instances not in 'running' state
      - <Region>     : One tab per region

.PARAMETER ConfigFile
    Path to the JSON configuration file. Defaults to config/aws.json.

.PARAMETER OutputDir
    Directory where the inventory workbook is written. Created if it does not exist.

.PARAMETER ArchiveDir
    Directory where the previous workbook is moved before a new run.

.PARAMETER TranscriptDir
    Directory for transcript log files.

.EXAMPLE
    .\Get-AllAWSInventory.ps1
    Runs with default paths relative to the script directory.

.EXAMPLE
    .\Get-AllAWSInventory.ps1 -OutputDir "D:\Reports\AWSInventory" -Verbose
    Runs with a custom output directory and verbose logging.
#>

[CmdletBinding()]
param(
    [Parameter()]
    [string]$ConfigFile = (Join-Path $PSScriptRoot 'config\aws.json'),

    [Parameter()]
    [string]$OutputDir = (Join-Path $PSScriptRoot 'Output\AWSInventory'),

    [Parameter()]
    [string]$ArchiveDir = (Join-Path $PSScriptRoot 'Output\AWSInventory\Archive'),

    [Parameter()]
    [string]$TranscriptDir = (Join-Path $PSScriptRoot 'Output\Transcripts')
)

$ErrorActionPreference = 'Stop'

#region Functions

function Backup-PreviousReport {
    <#
    .SYNOPSIS
        Archives a previous report file before a new run.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$SourcePath,

        [Parameter(Mandatory)]
        [string]$ArchivePath
    )

    if (-not (Test-Path -Path $SourcePath)) {
        Write-Verbose "No previous report to archive: $SourcePath"
        return
    }

    if (Test-Path -Path $ArchivePath) {
        Write-Verbose "Removing old archive: $ArchivePath"
        Remove-Item -Path $ArchivePath -Force
    }

    Write-Verbose "Archiving $SourcePath -> $ArchivePath"
    try {
        [System.IO.File]::Copy($SourcePath, $ArchivePath, $true)
        [System.IO.File]::Delete($SourcePath)
    }
    catch {
        Write-Warning "Could not archive previous report (OneDrive may have locked it): $_"
        Write-Warning "Deleting previous report without archiving: $SourcePath"
        [System.IO.File]::Delete($SourcePath)
    }
}

#endregion Functions

#region Initialization

# Ensure output directories exist
foreach ($dir in @($OutputDir, $ArchiveDir, $TranscriptDir)) {
    if (-not (Test-Path -Path $dir)) {
        New-Item -Path $dir -ItemType Directory -Force | Out-Null
        Write-Verbose "Created directory: $dir"
    }
}

# Start transcript
$transcriptPath = Join-Path $TranscriptDir "Get-AllAWSInventory_$(Get-Date -Format 'yyyy-MM-dd_HHmmss').log"
Start-Transcript -Path $transcriptPath

$startTime = Get-Date
Write-Host "Start of Get-AllAWSInventory.ps1 at $($startTime.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor Cyan

# Load configuration
if (-not (Test-Path -Path $ConfigFile)) {
    Write-Error "Configuration file not found: $ConfigFile"
    Stop-Transcript
    return
}

$config = Get-Content -Path $ConfigFile -Raw | ConvertFrom-Json
$credDir = Join-Path $PSScriptRoot $config.CredentialDir
Write-Host "Loaded $($config.Regions.Count) region(s) from config." -ForegroundColor Cyan

# Output file paths
$reportFile = Join-Path $OutputDir 'AWSInventory_All.xlsm'
$archiveFile = Join-Path $ArchiveDir 'AWSInventory_All.xlsm'

# Archive previous report
Backup-PreviousReport -SourcePath $reportFile -ArchivePath $archiveFile

#endregion Initialization

#region Collection Phase

$allInventoryData = [System.Collections.Generic.List[PSCustomObject]]::new()
$perRegionData = [ordered]@{}

$successCount = 0
$failCount = 0

# Load AWS credentials
$credPath = Join-Path $credDir $config.CredentialFile
if (-not (Test-Path -Path $credPath)) {
    Write-Error "Credential file not found: $credPath. Run Initialize-AWSCredentials.ps1 first."
    Stop-Transcript
    return
}

try {
    Write-Host "Loading AWS credentials..." -ForegroundColor Gray
    $credential = Import-Clixml -Path $credPath -ErrorAction Stop
    Set-AWSCredential -AccessKey $credential.UserName `
        -SecretKey $credential.GetNetworkCredential().Password
    Write-Host "AWS credentials loaded." -ForegroundColor Green
}
catch {
    Write-Error "Failed to load AWS credentials: $_"
    Stop-Transcript
    return
}

try {
    foreach ($region in $config.Regions) {
        $regionName = $region.Name
        $regionAlias = if ($region.Alias) { $region.Alias } else { $regionName }
        Write-Host "`nProcessing Region: $regionAlias ($regionName)" -ForegroundColor Cyan

        try {
            Set-DefaultAWSRegion -Region $regionName

            $allInstances = @(Get-EC2Instance -ErrorAction Stop |
                Select-Object -ExpandProperty Instances |
                Sort-Object { ($_.Tags | Where-Object { $_.Key -eq 'Name' }).Value })
            Write-Verbose "  Found $($allInstances.Count) instance(s) in $regionName"

            $totalInstances = $allInstances.Count
            $currentNum = 0

            $inventoryData = foreach ($instance in $allInstances) {
                $currentNum++
                $instanceName = ($instance.Tags | Where-Object { $_.Key -eq 'Name' }).Value
                Write-Progress -Activity "Scanning $regionAlias" `
                    -Status "$currentNum of $totalInstances - $instanceName" `
                    -PercentComplete ($currentNum / [math]::Max($totalInstances, 1) * 100)

                # --- Initialize all per-instance variables to prevent stale data ---
                $dnsName = ''
                $instanceStatus = 'unknown'

                # --- Instance basics (no redundant Get-EC2Instance call) ---
                $instanceId = $instance.InstanceId
                $instanceType = $instance.InstanceType.Value
                $instanceImageId = $instance.ImageId
                $instancePlatform = $instance.PlatformDetails
                $ipAddress = $instance.PrivateIpAddress
                $availabilityZone = $instance.Placement.AvailabilityZone

                # --- Instance status (use State from existing data, no extra API call) ---
                $instanceStatus = $instance.State.Name.Value
                if ([string]::IsNullOrEmpty($instanceStatus)) {
                    $instanceStatus = 'unknown'
                }

                # --- DNS name (reverse lookup with error handling) ---
                if (-not [string]::IsNullOrEmpty($ipAddress)) {
                    try {
                        $dnsResults = Resolve-DnsName -Name $ipAddress -DnsOnly -ErrorAction SilentlyContinue
                        if ($dnsResults) {
                            $dnsName = $dnsResults.NameHost
                            # If multiple DNS entries, prefer one starting with "AW" and longer than 3 chars
                            if ($dnsName -is [array] -and $dnsName.Count -gt 1) {
                                $filtered = $dnsName | Where-Object {
                                    $_ -like 'AW*' -and $_.Split('.')[0].Length -gt 3
                                }
                                $dnsName = if ($filtered) {
                                    if ($filtered -is [array]) { $filtered[0] } else { $filtered }
                                }
                                else {
                                    $dnsName[0]
                                }
                            }
                        }
                    }
                    catch {
                        Write-Verbose "    DNS lookup failed for $ipAddress : $_"
                    }
                }

                # --- Tags (null-safe) ---
                $tags = $instance.Tags
                $tagName       = ($tags | Where-Object { $_.Key -eq 'Name' }).Value
                $tagDesc       = ($tags | Where-Object { $_.Key -eq 'Description' }).Value
                $tagOS         = ($tags | Where-Object { $_.Key -eq 'OS' }).Value
                $tagBackup     = ($tags | Where-Object { $_.Key -eq 'Backup' }).Value
                $tagTeam       = ($tags | Where-Object { $_.Key -eq 'Team' }).Value
                $tagECR        = ($tags | Where-Object { $_.Key -eq 'ChangeRequest' }).Value
                $tagBuilder    = ($tags | Where-Object { $_.Key -eq 'Builder' }).Value
                $tagAMI        = ($tags | Where-Object { $_.Key -eq 'AMI' }).Value
                $tagConfigTest = ($tags | Where-Object { $_.Key -eq 'configtest' }).Value

                # --- Security groups (join names for Excel) ---
                $securityGroups = ($instance.SecurityGroups.GroupName | Where-Object { $_ }) -join '; '

                # --- CPU data ---
                $cpuCores = $instance.CpuOptions.CoreCount
                $cpuThreads = $instance.CpuOptions.ThreadsPerCore

                # --- Build the ordered property hashtable (20 columns) ---
                [PSCustomObject]([ordered]@{
                    'AWS_Instance'     = $instanceId
                    'DNS_Name'         = $dnsName
                    'Tag_Name'         = $tagName
                    'Tag_Description'  = $tagDesc
                    'Tag_OS'           = $tagOS
                    'Tag_Backup'       = $tagBackup
                    'AWS_Type'         = $instanceType
                    'Status'           = $instanceStatus
                    'Image'            = $instanceImageId
                    'Platform'         = $instancePlatform
                    'IPAddress'        = $ipAddress
                    'Security_Group'   = $securityGroups
                    'CPU_Cores'        = $cpuCores
                    'CPU_Threads'      = $cpuThreads
                    'Availability_Zone' = $availabilityZone
                    'Tag_Team'         = $tagTeam
                    'Tag_ECR'          = $tagECR
                    'Tag_Builder'      = $tagBuilder
                    'Tag_AMI'          = $tagAMI
                    'Tag_ConfigTest'   = $tagConfigTest
                })
            }

            Write-Progress -Activity "Scanning $regionAlias" -Completed

            # Store collected data
            $regionInventory = @($inventoryData)
            foreach ($item in $regionInventory) {
                $allInventoryData.Add($item)
            }
            $perRegionData[$regionAlias] = $regionInventory

            Write-Host "  Collected $($regionInventory.Count) instance(s)." -ForegroundColor Green
            $successCount++
        }
        catch {
            Write-Warning "Failed to process region '$regionName': $_"
            $failCount++
        }
    }
}
finally {
    Write-Verbose "Clearing AWS credentials"
    Clear-AWSCredential -ErrorAction SilentlyContinue
    Clear-DefaultAWSRegion -ErrorAction SilentlyContinue
}

#endregion Collection Phase

#region Build Workbook

if ($allInventoryData.Count -eq 0) {
    Write-Warning "No instance data collected from any region. Skipping workbook creation."
    Stop-Transcript
    return
}

Write-Host "`nBuilding workbook with $($allInventoryData.Count) total instance(s)..." -ForegroundColor Cyan

# Build filtered views
$notRunningData = @($allInventoryData | Where-Object { $_.Status -ne 'running' })

# Conditional formatting rules
$awsCfRules = @(
    # Gray: Not running instances
    New-ConditionalText -Text 'stopped' -Range 'H:H' -BackgroundColor LightGray
    New-ConditionalText -Text 'terminated' -Range 'H:H' -BackgroundColor LightGray
    New-ConditionalText -Text 'shutting-down' -Range 'H:H' -BackgroundColor LightGray
    New-ConditionalText -Text 'stopping' -Range 'H:H' -BackgroundColor LightGray
)

# Export All_VMs tab first (this creates the workbook)
$tempXlsx = Join-Path $OutputDir 'AWSInventory_All.tmp.xlsm'
if (Test-Path $tempXlsx) { Remove-Item $tempXlsx -Force }
if (Test-Path $reportFile) { Remove-Item $reportFile -Force }

$allInventoryData | Export-Excel -Path $tempXlsx -WorksheetName 'All_VMs' `
    -AutoSize -FreezePane 2, 2 -BoldTopRow -ConditionalText $awsCfRules `
    -TableName 'All_VMs' -TableStyle Medium9

# Not_Running
if ($notRunningData.Count -gt 0) {
    $notRunningData | Export-Excel -Path $tempXlsx -WorksheetName 'Not_Running' `
        -AutoSize -FreezePane 2, 2 -BoldTopRow -ConditionalText $awsCfRules `
        -TableName 'Not_Running' -TableStyle Medium9
}
else {
    Export-Excel -Path $tempXlsx -WorksheetName 'Not_Running' -InputObject $null
}

# Per-region tabs
foreach ($regionAlias in $perRegionData.Keys) {
    $regionData = $perRegionData[$regionAlias]
    # Sanitize region name for worksheet name (max 31 chars, no special chars)
    $tabName = $regionAlias -replace '[:\\/\?\*\[\]]', '_'
    if ($tabName.Length -gt 31) { $tabName = $tabName.Substring(0, 31) }

    if ($regionData.Count -gt 0) {
        $regionData | Export-Excel -Path $tempXlsx -WorksheetName $tabName `
            -AutoSize -FreezePane 2, 2 -BoldTopRow -ConditionalText $awsCfRules `
            -TableName ($tabName -replace '[^A-Za-z0-9_]', '_') -TableStyle Medium9
    }
    else {
        Export-Excel -Path $tempXlsx -WorksheetName $tabName -InputObject $null
    }
}

#endregion Build Workbook

#region Search Tab + VBA

# Open the workbook and add the Search tab with VBA
$pkg = Open-ExcelPackage $tempXlsx

# Resolve EPPlus enum values at runtime
$epAsm = $pkg.GetType().Assembly
$epTypes = $epAsm.GetExportedTypes()
$findType = { param([string]$Name) $epTypes | Where-Object { $_.Name -eq $Name } | Select-Object -First 1 }

$borderThin     = [Enum]::Parse((& $findType 'ExcelBorderStyle'), 'Thin')
$fillSolid      = [Enum]::Parse((& $findType 'ExcelFillStyle'), 'Solid')
$shapeRoundRect = [Enum]::Parse((& $findType 'eShapeStyle'), 'RoundRect')
$textCenter     = [Enum]::Parse((& $findType 'eTextAlignment'), 'Center')
$fillSolidFill  = [Enum]::Parse((& $findType 'eFillStyle'), 'SolidFill')

$dataWs = $pkg.Workbook.Worksheets['All_VMs']
$searchWs = $pkg.Workbook.Worksheets.Add('Search')
$pkg.Workbook.Worksheets.MoveToStart('Search')

# Build the Search tab layout
$searchWs.Cells[1, 1].Value = 'AWS EC2 Inventory - Search'
$searchWs.Cells[1, 1].Style.Font.Bold = $true
$searchWs.Cells[1, 1].Style.Font.Size = 16
$searchWs.Cells[1, 1].Style.Font.Color.SetColor([System.Drawing.Color]::FromArgb(68, 114, 196))

$searchWs.Cells[3, 1].Value = 'Enter search term:'
$searchWs.Cells[3, 1].Style.Font.Bold = $true
$searchWs.Cells[3, 1].Style.Font.Size = 11

# Style the search input cell B3
$searchWs.Cells[3, 2].Style.Font.Size = 11
$searchWs.Cells[3, 2].Style.Border.Top.Style = $borderThin
$searchWs.Cells[3, 2].Style.Border.Bottom.Style = $borderThin
$searchWs.Cells[3, 2].Style.Border.Left.Style = $borderThin
$searchWs.Cells[3, 2].Style.Border.Right.Style = $borderThin
$searchWs.Column(2).Width = 40

# Column widths
$searchWs.Column(1).Width = 20

# Copy headers to row 5
$lastCol = $dataWs.Dimension.End.Column
for ($c = 1; $c -le $lastCol; $c++) {
    $searchWs.Cells[5, $c].Value = $dataWs.Cells[1, $c].Value
    $searchWs.Cells[5, $c].Style.Font.Bold = $true
    $searchWs.Cells[5, $c].Style.Fill.PatternType = $fillSolid
    $searchWs.Cells[5, $c].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(68, 114, 196))
    $searchWs.Cells[5, $c].Style.Font.Color.SetColor([System.Drawing.Color]::White)
}

# Freeze panes below header row
$searchWs.View.FreezePanes(6, 2)

# Add Go button (search)
$goButton = $searchWs.Drawings.AddShape('GoButton', $shapeRoundRect)
$goButton.SetPosition(2, 0, 2, 0)
$goButton.SetSize(80, 30)
$goButton.Text = 'Search'
$goButton.TextAlignment = $textCenter
$goButton.Fill.Style = $fillSolidFill
$goButton.Fill.Color = [System.Drawing.Color]::FromArgb(68, 114, 196)
$goButton.Font.Color = [System.Drawing.Color]::White
$goButton.Font.Bold = $true
$goButton.Font.Size = 11

# VBA project
$pkg.Workbook.CreateVBAProject()

# Assign macros via Workbook_Open
$pkg.Workbook.CodeModule.Code = @"
Private Sub Workbook_Open()
    ThisWorkbook.Worksheets("Search").Shapes("GoButton").OnAction = "RunSearch"
End Sub
"@

# Add VBA module with search logic
$vbaModule = $pkg.Workbook.VbaProject.Modules.AddModule('SearchModule')
$vbaModule.Code = @"
Public Sub RunSearch()
    Dim searchWs As Worksheet
    Dim dataWs As Worksheet
    Set searchWs = ThisWorkbook.Worksheets("Search")
    Set dataWs = ThisWorkbook.Worksheets("All_VMs")

    Dim searchVal As String
    searchVal = LCase(Trim(searchWs.Range("B3").Value))

    Application.ScreenUpdating = False

    ' Clear previous search results (row 6 and below, keep rows 1-5)
    Dim lastResultRow As Long
    lastResultRow = searchWs.Cells(searchWs.Rows.Count, 1).End(xlUp).Row
    If lastResultRow > 5 Then
        searchWs.Rows("6:" & lastResultRow).Delete
    End If

    If searchVal = "" Then
        MsgBox "Please enter a search term.", vbInformation, "Search"
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' Search the data table
    Dim tbl As ListObject
    Set tbl = dataWs.ListObjects("All_VMs")
    Dim lastCol As Long
    lastCol = tbl.ListColumns.Count

    Dim resultRow As Long
    resultRow = 6
    Dim matchCount As Long
    matchCount = 0

    Dim r As Long
    For r = 1 To tbl.DataBodyRange.Rows.Count
        Dim matched As Boolean
        matched = False
        Dim c As Long
        For c = 1 To lastCol
            If InStr(1, LCase(CStr(tbl.DataBodyRange.Cells(r, c).Value)), searchVal, vbTextCompare) > 0 Then
                matched = True
                Exit For
            End If
        Next c
        If matched Then
            Dim col As Long
            For col = 1 To lastCol
                searchWs.Cells(resultRow, col).Value = tbl.DataBodyRange.Cells(r, col).Value
            Next col
            resultRow = resultRow + 1
            matchCount = matchCount + 1
        End If
    Next r

    ' Auto-fit columns on search sheet
    searchWs.Columns("A:" & Chr(64 + Application.WorksheetFunction.Min(lastCol, 26))).AutoFit

    ' Status message
    searchWs.Cells(4, 1).Value = matchCount & " result(s) found"
    searchWs.Cells(4, 1).Font.Italic = True
    searchWs.Cells(4, 1).Font.Color = RGB(100, 100, 100)

    Application.ScreenUpdating = True
End Sub
"@

Close-ExcelPackage $pkg -SaveAs $reportFile
$pkg = $null
Remove-Item $tempXlsx -Force -ErrorAction SilentlyContinue

Write-Host "  Workbook saved: $reportFile" -ForegroundColor Green

# Report tab summary
Write-Host "`n  Tabs created:" -ForegroundColor Cyan
Write-Host "    Search           : Search UI" -ForegroundColor Gray
Write-Host "    All_VMs          : $($allInventoryData.Count) instance(s)" -ForegroundColor Gray
Write-Host "    Not_Running      : $($notRunningData.Count) instance(s)" -ForegroundColor Gray
foreach ($regionAlias in $perRegionData.Keys) {
    Write-Host "    $regionAlias : $($perRegionData[$regionAlias].Count) instance(s)" -ForegroundColor Gray
}

#endregion Search Tab + VBA

#region Summary

$endTime = Get-Date
$duration = $endTime - $startTime
Write-Host "`n--- Summary ---" -ForegroundColor Cyan
Write-Host "  Succeeded: $successCount" -ForegroundColor Green
if ($failCount -gt 0) {
    Write-Host "  Failed:    $failCount" -ForegroundColor Red
}
Write-Host "  Duration:  $($duration.ToString('hh\:mm\:ss'))"
Write-Host "  Transcript: $transcriptPath"
Write-Host "End of Get-AllAWSInventory.ps1 at $($endTime.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor Cyan

Stop-Transcript

#endregion Summary
