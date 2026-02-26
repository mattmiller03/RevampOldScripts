#Requires -Modules Microsoft.PowerShell.SecretManagement, ImportExcel

<#
.SYNOPSIS
    Collects VM inventory from all vCenters defined in the configuration file.

.DESCRIPTION
    Iterates through each vCenter server listed in config/vcenters.json, connects using
    credentials stored in a SecretManagement vault, collects VM inventory, and saves
    the results as per-vCenter CSV files.

    Previous inventory files are archived before new ones are created.

    Credentials must be set up first by running Initialize-VCenterSecrets.ps1.

.PARAMETER ConfigFile
    Path to the JSON configuration file. Defaults to config/vcenters.json.

.PARAMETER OutputDir
    Directory where inventory CSV files are written. Created if it does not exist.

.PARAMETER ArchiveDir
    Directory where previous inventory files are moved before a new run. Created if it does not exist.

.PARAMETER TranscriptDir
    Directory for transcript log files. Created if it does not exist.

.EXAMPLE
    .\Get-AllVMInventory.ps1
    Runs with default paths relative to the script directory.

.EXAMPLE
    .\Get-AllVMInventory.ps1 -OutputDir "D:\Reports\VMInventory" -Verbose
    Runs with a custom output directory and verbose logging.
#>

[CmdletBinding()]
param(
    [Parameter()]
    [string]$ConfigFile = (Join-Path $PSScriptRoot 'config\vcenters.json'),

    [Parameter()]
    [string]$OutputDir = (Join-Path $PSScriptRoot 'Output\VMInventory'),

    [Parameter()]
    [string]$ArchiveDir = (Join-Path $PSScriptRoot 'Output\VMInventory\Archive'),

    [Parameter()]
    [string]$TranscriptDir = (Join-Path $PSScriptRoot 'Output\Transcripts')
)

$ErrorActionPreference = 'Stop'

# Accept either VMware.PowerCLI or VCF.PowerCLI (Broadcom rebrand)
$powerCLI = Get-Module -ListAvailable -Name 'VCF.PowerCLI', 'VMware.PowerCLI' | Select-Object -First 1
if (-not $powerCLI) {
    Write-Error "Neither VMware.PowerCLI nor VCF.PowerCLI is installed. Install one of them to continue."
    return
}
Import-Module $powerCLI.Name -ErrorAction Stop
Write-Verbose "Loaded $($powerCLI.Name) v$($powerCLI.Version)"

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

    # Remove existing archive copy if present
    if (Test-Path -Path $ArchivePath) {
        Write-Verbose "Removing old archive: $ArchivePath"
        Remove-Item -Path $ArchivePath -Force
    }

    Write-Verbose "Archiving $SourcePath -> $ArchivePath"
    Move-Item -Path $SourcePath -Destination $ArchivePath -Force
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
$transcriptPath = Join-Path $TranscriptDir "Get-AllVMInventory_$(Get-Date -Format 'yyyy-MM-dd_HHmmss').log"
Start-Transcript -Path $transcriptPath

$startTime = Get-Date
Write-Host "Start of Get-AllVMInventory.ps1 at $($startTime.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor Cyan

# Load configuration
if (-not (Test-Path -Path $ConfigFile)) {
    Write-Error "Configuration file not found: $ConfigFile"
    Stop-Transcript
    return
}

$config = Get-Content -Path $ConfigFile -Raw | ConvertFrom-Json
$vaultName = $config.VaultName
Write-Host "Loaded $($config.VCenters.Count) vCenter(s) from config." -ForegroundColor Cyan

#endregion Initialization

#region Main Loop

$successCount = 0
$failCount = 0

foreach ($vc in $config.VCenters) {
    $vcName = $vc.Name
    Write-Host "`nProcessing vCenter: $vcName" -ForegroundColor Cyan

    $reportFile = Join-Path $OutputDir "VMInventory_$vcName.xlsm"
    $archiveFile = Join-Path $ArchiveDir "VMInventory_$vcName.xlsm"

    # Archive previous report
    Backup-PreviousReport -SourcePath $reportFile -ArchivePath $archiveFile

    $connection = $null
    try {
        # Retrieve credential from SecretManagement vault
        Write-Verbose "Retrieving credential for '$($vc.SecretName)' from vault '$vaultName'"
        $credential = Get-Secret -Name $vc.SecretName -Vault $vaultName -ErrorAction Stop

        # Connect to vCenter
        Write-Host "  Connecting to $vcName..." -ForegroundColor Gray
        $connection = Connect-VIServer -Server $vcName -Credential $credential -ErrorAction Stop
        Write-Host "  Connected to $vcName." -ForegroundColor Green

        # Collect VM inventory
        Write-Host "  Collecting VM inventory..." -ForegroundColor Gray
        $allVMs = Get-VM -Server $connection -ErrorAction Stop
        Write-Verbose "  Found $($allVMs.Count) VM(s) on $vcName"

        $inventoryData = foreach ($vm in $allVMs) {
            Write-Verbose "    Processing VM: $($vm.Name)"
            $vmView = Get-View -VIObject $vm -Property Config, Guest, Runtime, Summary

            # Guest OS
            $configuredOS = $vmView.Config.GuestFullName
            $runningOS = $vmView.Guest.GuestFullName

            # Devices from config
            $devices = $vmView.Config.Hardware.Device
            $floppyDrive = $null -ne ($devices | Where-Object { $_ -is [VMware.Vim.VirtualFloppy] })
            $usbControllers = @($devices | Where-Object { $_ -is [VMware.Vim.VirtualUSBController] -or $_ -is [VMware.Vim.VirtualUSBXHCIController] })
            $hasUSB = $usbControllers.Count -gt 0
            $usbCount = $usbControllers.Count

            # CD drive
            $cdDrive = $devices | Where-Object { $_ -is [VMware.Vim.VirtualCdrom] } | Select-Object -First 1
            $cdDrivePresent = $null -ne $cdDrive
            $cdConnected = if ($cdDrive) { $cdDrive.Connectable.Connected } else { $false }

            # Disk consolidation
            $needsConsolidation = $vmView.Runtime.ConsolidationNeeded

            # CPU and memory hot-add
            $cpuHotAdd = $vmView.Config.CpuHotAddEnabled
            $memHotAdd = $vmView.Config.MemoryHotAddEnabled

            # Hardware version
            $hwVersion = $vmView.Config.Version

            # VMTools
            $toolsStatus = $vmView.Guest.ToolsStatus
            $toolsVersion = $vmView.Guest.ToolsVersion

            # Template
            $isTemplate = $vm.ExtensionData.Config.Template

            # IP addresses (up to 3)
            $guestNics = $vmView.Guest.Net
            $ips = @('')*3
            $ipIndex = 0
            if ($guestNics) {
                foreach ($gNic in $guestNics) {
                    if ($gNic.IpAddress) {
                        foreach ($addr in $gNic.IpAddress) {
                            if ($ipIndex -lt 3 -and $addr -match '^\d+\.\d+\.\d+\.\d+$') {
                                $ips[$ipIndex] = $addr
                                $ipIndex++
                            }
                        }
                    }
                }
            }

            # Network adapters (up to 5)
            $netAdapters = Get-NetworkAdapter -VM $vm -ErrorAction SilentlyContinue | Sort-Object Name
            $nics = @('')*5
            for ($i = 0; $i -lt [math]::Min($netAdapters.Count, 5); $i++) {
                $nics[$i] = $netAdapters[$i].NetworkName
            }

            # Hard disks (up to 12)
            $hardDisks = Get-HardDisk -VM $vm -ErrorAction SilentlyContinue | Sort-Object Name
            $disks = @('')*12
            for ($i = 0; $i -lt [math]::Min($hardDisks.Count, 12); $i++) {
                $disks[$i] = [math]::Round($hardDisks[$i].CapacityGB, 2)
            }

            # Firmware and security
            $firmware = $vmView.Config.Firmware
            $nestedHV = $vmView.Config.NestedHVEnabled
            $efiSecureBoot = $vmView.Config.BootOptions.EfiSecureBootEnabled
            $vvtd = $vmView.Config.VvtdEnabled
            $vbs = $vmView.Config.VbsEnabled

            # Hostname vs VM name comparison
            $guestHostname = $vmView.Guest.HostName
            $nameNotEqual = if ($guestHostname -and $guestHostname -ne $vm.Name) { $true } else { $false }

            # vTPM
            $vtpm = $null -ne ($devices | Where-Object { $_ -is [VMware.Vim.VirtualTPM] })

            # Tags (by category)
            $tagAssignments = Get-TagAssignment -Entity $vm -ErrorAction SilentlyContinue
            $tagsByCategory = @{}
            if ($tagAssignments) {
                foreach ($ta in $tagAssignments) {
                    $catName = $ta.Tag.Category.Name
                    if (-not $tagsByCategory.ContainsKey($catName)) {
                        $tagsByCategory[$catName] = [System.Collections.Generic.List[string]]::new()
                    }
                    $tagsByCategory[$catName].Add($ta.Tag.Name)
                }
            }

            # Helper to get tag values by category (supports multiple values)
            $getTag = {
                param([string]$Category, [int]$Index)
                if ($tagsByCategory.ContainsKey($Category) -and $tagsByCategory[$Category].Count -gt $Index) {
                    return $tagsByCategory[$Category][$Index]
                }
                return ''
            }

            # Custom attributes
            $customAttribs = Get-Annotation -Entity $vm -ErrorAction SilentlyContinue
            $changeNumber = ($customAttribs | Where-Object { $_.Name -eq 'Chg-number' }).Value
            $lastBackup = ($customAttribs | Where-Object { $_.Name -eq 'NB_Last_Backup' }).Value

            # Cluster, Datacenter, ResourcePool, Folder
            $vmHost = $vm.VMHost
            $vmCluster = Get-Cluster -VMHost $vmHost -ErrorAction SilentlyContinue
            $vmDatacenter = Get-Datacenter -VM $vm -ErrorAction SilentlyContinue

            [PSCustomObject]@{
                'Name'                    = $vm.Name
                'PowerState'              = $vm.PowerState
                'Cluster'                 = if ($vmCluster) { $vmCluster.Name } else { '' }
                'Configured_Guest_OS'     = $configuredOS
                'Running_Guest_OS'        = $runningOS
                'Notes'                   = $vm.Notes
                'Floppydrive'             = $floppyDrive
                'USB Controller'          = $hasUSB
                'No USB Cont'             = $usbCount
                'Needs Disk Consolidated' = $needsConsolidation
                'NumCpu'                  = $vm.NumCpu
                'CPU Hot Add'             = $cpuHotAdd
                'Memory GB'               = [math]::Round($vm.MemoryGB, 2)
                'Memory Hot Add'          = $memHotAdd
                'Hardware Version'        = $hwVersion
                'VMTools Status'          = $toolsStatus
                'Datacenter'              = if ($vmDatacenter) { $vmDatacenter.Name } else { '' }
                'CD Drive'                = $cdDrivePresent
                'CD Connected'            = $cdConnected
                'Template'                = $isTemplate
                'IP #1'                   = $ips[0]
                'IP #2'                   = $ips[1]
                'IP #3'                   = $ips[2]
                '1st vNic'                = $nics[0]
                '2nd vNic'                = $nics[1]
                '3rd vNic'                = $nics[2]
                '4th vNic'                = $nics[3]
                '5th vNic'                = $nics[4]
                'VMTools Version'         = $toolsVersion
                'Disk1'                   = $disks[0]
                'Disk2'                   = $disks[1]
                'Disk3'                   = $disks[2]
                'Disk4'                   = $disks[3]
                'Disk5'                   = $disks[4]
                'Disk6'                   = $disks[5]
                'Disk7'                   = $disks[6]
                'Disk8'                   = $disks[7]
                'Disk9'                   = $disks[8]
                'Disk10'                  = $disks[9]
                'Disk11'                  = $disks[10]
                'Disk12'                  = $disks[11]
                'vCenter'                 = $vcName
                'Firmware'                = $firmware
                'NestedHVEnabled'         = $nestedHV
                'EfiSecureBootEnabled'    = $efiSecureBoot
                'VvtdEnabled'             = $vvtd
                'VbsEnabled'              = $vbs
                'Hostname not equal VMname' = $nameNotEqual
                'Application_Tag1'        = & $getTag 'Application' 0
                'Application_Tag2'        = & $getTag 'Application' 1
                'Function_Tag'            = & $getTag 'Function' 0
                'OperatingSystem_Tag'     = & $getTag 'OperatingSystem' 0
                'EnclaveID_Tag'           = & $getTag 'EnclaveID' 0
                'ResourceType_Tag'        = & $getTag 'ResourceType' 0
                'Site_Location_Tag'       = & $getTag 'Site_Location' 0
                'VlanID_Tag1'             = & $getTag 'VlanID' 0
                'VlanID_Tag2'             = & $getTag 'VlanID' 1
                'VlanID_Tag3'             = & $getTag 'VlanID' 2
                'VlanID_Tag4'             = & $getTag 'VlanID' 3
                'TeamPoc_Tag1'            = & $getTag 'TeamPoc' 0
                'TeamPoc_Tag2'            = & $getTag 'TeamPoc' 1
                'TeamPoc_Tag3'            = & $getTag 'TeamPoc' 2
                'TeamPoc_Tag4'            = & $getTag 'TeamPoc' 3
                'Change_Number'           = $changeNumber
                'vTPM'                    = $vtpm
                'ResourcePool'            = if ($vm.ResourcePool) { $vm.ResourcePool.Name } else { '' }
                'FolderName'              = if ($vm.Folder) {
                    $folderPath = @()
                    $f = $vm.Folder
                    while ($f -and $f.Name -ne 'vm') {
                        $folderPath += $f.Name
                        $f = $f.Parent
                    }
                    [array]::Reverse($folderPath)
                    $folderPath -join '/'
                } else { '' }
                'LastBackup'              = $lastBackup
            }
        }

        # Conditional formatting rules
        $vmCfRules = @(
            # Gray: Powered off VMs
            New-ConditionalText -Text 'PoweredOff' -Range 'B:B' -BackgroundColor LightGray
            # Red: Needs disk consolidation
            New-ConditionalText -Text 'True' -Range 'J:J' -BackgroundColor Red -ConditionalTextColor White
            # Yellow: VMTools not running or not installed
            New-ConditionalText -Text 'toolsNotRunning' -Range 'P:P' -BackgroundColor Yellow
            New-ConditionalText -Text 'toolsNotInstalled' -Range 'P:P' -BackgroundColor Red -ConditionalTextColor White
            New-ConditionalText -Text 'toolsOld' -Range 'P:P' -BackgroundColor Yellow
            # Yellow: CD connected
            New-ConditionalText -Text 'True' -Range 'S:S' -BackgroundColor Yellow
            # Yellow: Floppy drive present
            New-ConditionalText -Text 'True' -Range 'G:G' -BackgroundColor Yellow
            # Yellow: Old hardware versions (vmx-13 and below = pre-vSphere 6.7)
            New-ConditionalText -Text 'vmx-07' -Range 'O:O' -BackgroundColor Orange -ConditionalTextColor White
            New-ConditionalText -Text 'vmx-08' -Range 'O:O' -BackgroundColor Orange -ConditionalTextColor White
            New-ConditionalText -Text 'vmx-09' -Range 'O:O' -BackgroundColor Orange -ConditionalTextColor White
            New-ConditionalText -Text 'vmx-10' -Range 'O:O' -BackgroundColor Yellow
            New-ConditionalText -Text 'vmx-11' -Range 'O:O' -BackgroundColor Yellow
            New-ConditionalText -Text 'vmx-12' -Range 'O:O' -BackgroundColor Yellow
            New-ConditionalText -Text 'vmx-13' -Range 'O:O' -BackgroundColor Yellow
        )

        # Export data to temp xlsx, then add Search tab and save as xlsm
        $tempXlsx = Join-Path $OutputDir "VMInventory_$vcName.tmp.xlsx"
        if (Test-Path $tempXlsx) { Remove-Item $tempXlsx -Force }
        if (Test-Path $reportFile) { Remove-Item $reportFile -Force }

        $inventoryData | Export-Excel -Path $tempXlsx -WorksheetName 'VMInventory' `
            -AutoSize -FreezeTopRow -BoldTopRow -ConditionalText $vmCfRules `
            -TableName 'VMInventory' -TableStyle Medium6

        # Add Search tab with input box and Go button
        $pkg = Open-ExcelPackage $tempXlsx

        # Resolve EPPlus enum values at runtime (search by short name to handle namespace changes across versions)
        $epAsm = $pkg.GetType().Assembly
        $epTypes = $epAsm.GetExportedTypes()
        $findType = { param([string]$Name) $epTypes | Where-Object { $_.Name -eq $Name } | Select-Object -First 1 }

        $borderThin     = [Enum]::Parse((& $findType 'ExcelBorderStyle'), 'Thin')
        $fillSolid      = [Enum]::Parse((& $findType 'ExcelFillStyle'), 'Solid')
        $shapeRoundRect = [Enum]::Parse((& $findType 'eShapeStyle'), 'RoundRect')
        $textCenter     = [Enum]::Parse((& $findType 'eTextAlignment'), 'Center')
        $fillSolidFill  = [Enum]::Parse((& $findType 'eFillStyle'), 'SolidFill')
        $dataWs = $pkg.Workbook.Worksheets['VMInventory']
        $searchWs = $pkg.Workbook.Worksheets.Add('Search')
        $pkg.Workbook.Worksheets.MoveToStart('Search')

        # Build the Search tab layout
        $searchWs.Cells[1, 1].Value = 'VM Inventory - Search & New Entry'
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

        # Row 6: New entry input row (light green background with borders)
        $lightGreen = [System.Drawing.Color]::FromArgb(226, 239, 218)
        for ($c = 1; $c -le $lastCol; $c++) {
            $searchWs.Cells[6, $c].Style.Fill.PatternType = $fillSolid
            $searchWs.Cells[6, $c].Style.Fill.BackgroundColor.SetColor($lightGreen)
            $searchWs.Cells[6, $c].Style.Border.Top.Style = $borderThin
            $searchWs.Cells[6, $c].Style.Border.Bottom.Style = $borderThin
            $searchWs.Cells[6, $c].Style.Border.Left.Style = $borderThin
            $searchWs.Cells[6, $c].Style.Border.Right.Style = $borderThin
        }

        # Freeze panes below the input row
        $searchWs.View.FreezePanes(7, 1)

        # Row 7: separator label for search results
        $searchWs.Cells[7, 1].Value = 'Search Results:'
        $searchWs.Cells[7, 1].Style.Font.Bold = $true
        $searchWs.Cells[7, 1].Style.Font.Size = 10
        $searchWs.Cells[7, 1].Style.Font.Color.SetColor([System.Drawing.Color]::Gray)

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

        # Add Entry button (green, next to Go)
        $addButton = $searchWs.Drawings.AddShape('AddEntryButton', $shapeRoundRect)
        $addButton.SetPosition(2, 0, 3, 20)
        $addButton.SetSize(100, 30)
        $addButton.Text = 'Add Entry'
        $addButton.TextAlignment = $textCenter
        $addButton.Fill.Style = $fillSolidFill
        $addButton.Fill.Color = [System.Drawing.Color]::FromArgb(112, 173, 71)
        $addButton.Font.Color = [System.Drawing.Color]::White
        $addButton.Font.Bold = $true
        $addButton.Font.Size = 11

        # VBA project
        $pkg.Workbook.CreateVBAProject()

        # Assign macros via Workbook_Open (compatible with all EPPlus versions)
        $pkg.Workbook.CodeModule.Code = @"
Private Sub Workbook_Open()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Search")
    ws.Shapes("GoButton").OnAction = "RunSearch"
    ws.Shapes("AddEntryButton").OnAction = "AddEntry"
End Sub
"@

        # Add VBA module with search and add entry logic
        $vbaModule = $pkg.Workbook.VbaProject.Modules.AddModule('SearchModule')
        $vbaModule.Code = @"
Public Sub RunSearch()
    Dim searchWs As Worksheet
    Dim dataWs As Worksheet
    Set searchWs = ThisWorkbook.Worksheets("Search")
    Set dataWs = ThisWorkbook.Worksheets("VMInventory")

    Dim searchVal As String
    searchVal = LCase(Trim(searchWs.Range("B3").Value))

    Application.ScreenUpdating = False

    ' Clear previous search results (row 8 and below, keep rows 1-7)
    Dim lastResultRow As Long
    lastResultRow = searchWs.Cells(searchWs.Rows.Count, 1).End(xlUp).Row
    If lastResultRow > 7 Then
        searchWs.Rows("8:" & lastResultRow).Delete
    End If

    If searchVal = "" Then
        MsgBox "Please enter a search term.", vbInformation, "Search"
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' Search the data table
    Dim tbl As ListObject
    Set tbl = dataWs.ListObjects("VMInventory")
    Dim lastCol As Long
    lastCol = tbl.ListColumns.Count

    Dim resultRow As Long
    resultRow = 8
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

Public Sub AddEntry()
    Dim searchWs As Worksheet
    Dim dataWs As Worksheet
    Set searchWs = ThisWorkbook.Worksheets("Search")
    Set dataWs = ThisWorkbook.Worksheets("VMInventory")

    ' Check that at least the Name field (column 1) has a value
    If Trim(searchWs.Cells(6, 1).Value) = "" Then
        MsgBox "Please enter at least a VM Name in the first cell of the green row.", vbExclamation, "Add Entry"
        Exit Sub
    End If

    Application.ScreenUpdating = False

    ' Get the data table and add a new row
    Dim tbl As ListObject
    Set tbl = dataWs.ListObjects("VMInventory")
    Dim lastCol As Long
    lastCol = tbl.ListColumns.Count

    ' Add a new row to the table
    Dim newRow As ListRow
    Set newRow = tbl.ListRows.Add

    ' Copy values from input row 6 to the new table row
    Dim c As Long
    For c = 1 To lastCol
        If searchWs.Cells(6, c).Value <> "" Then
            newRow.Range.Cells(1, c).Value = searchWs.Cells(6, c).Value
        End If
    Next c

    ' Clear the input row
    searchWs.Range(searchWs.Cells(6, 1), searchWs.Cells(6, lastCol)).ClearContents

    Application.ScreenUpdating = True
    MsgBox "New VM entry added to the inventory.", vbInformation, "Add Entry"
End Sub
"@

        Close-ExcelPackage $pkg -SaveAs $reportFile
        Remove-Item $tempXlsx -Force -ErrorAction SilentlyContinue
        Write-Host "  Collected $(@($inventoryData).Count) VM(s)." -ForegroundColor Green
        $successCount++
    }
    catch {
        Write-Warning "Failed to process vCenter '$vcName': $_"
        $failCount++
    }
    finally {
        if ($null -ne $connection) {
            Write-Verbose "Disconnecting from $vcName"
            Disconnect-VIServer -Server $connection -Confirm:$false -ErrorAction SilentlyContinue
        }
    }
}

#endregion Main Loop

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
Write-Host "End of Get-AllVMInventory.ps1 at $($endTime.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor Cyan

Stop-Transcript

#endregion Summary
