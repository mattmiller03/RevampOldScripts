#Requires -Modules ImportExcel

<#
.SYNOPSIS
    Collects VM inventory from all vCenters and produces a multi-tab Excel workbook.

.DESCRIPTION
    Iterates through each vCenter server listed in config/vcenters.json, connects using
    DPAPI-encrypted credentials (created by Initialize-VCenterCredentials.ps1), collects
    VM inventory, and produces a single Excel workbook with multiple tabs:

      - Search      : Search UI with VBA macro (searches All_VMs table)
      - All_VMs     : Combined inventory from all vCenters
      - MissingTags : VMs missing any required tag category (configurable in JSON)
      - VM_BIOS     : VMs using BIOS firmware (not EFI)
      - VMs_Powered_Off : VMs in PoweredOff state
      - <vCenter>   : One tab per vCenter with that vCenter's VMs

    Tag categories and column counts are driven by the RequiredTags array in the JSON config.

.PARAMETER ConfigFile
    Path to the JSON configuration file. Defaults to config/vcenters.json.

.PARAMETER OutputDir
    Directory where the inventory workbook is written. Created if it does not exist.

.PARAMETER ArchiveDir
    Directory where the previous workbook is moved before a new run.

.PARAMETER TranscriptDir
    Directory for transcript log files.

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
    [string]$TranscriptDir = (Join-Path $PSScriptRoot 'Output\Transcripts'),

    [Parameter()]
    [switch]$SkipModuleCheck,

    [Parameter()]
    [pscredential]$VCenterCredential
)

$ErrorActionPreference = 'Stop'

# Accept either VMware.PowerCLI or VCF.PowerCLI (Broadcom rebrand)
if (-not $SkipModuleCheck) {
    $powerCLI = Get-Module -ListAvailable -Name 'VCF.PowerCLI', 'VMware.PowerCLI' | Select-Object -First 1
    if (-not $powerCLI) {
        Write-Error "Neither VMware.PowerCLI nor VCF.PowerCLI is installed. Install one of them to continue."
        return
    }
    Import-Module $powerCLI.Name -ErrorAction Stop
    Write-Verbose "Loaded $($powerCLI.Name) v$($powerCLI.Version)"
}
else {
    Write-Verbose "Skipping PowerCLI module check"
}

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
$credDir = Join-Path $PSScriptRoot $config.CredentialDir
Write-Host "Loaded $($config.VCenters.Count) vCenter(s) from config." -ForegroundColor Cyan
Write-Host "Tag categories: $($config.RequiredTags.Count) configured." -ForegroundColor Cyan

# Output file paths
$reportFile = Join-Path $OutputDir 'VMInventory_All.xlsm'
$archiveFile = Join-Path $ArchiveDir 'VMInventory_All.xlsm'

# Archive previous report
Backup-PreviousReport -SourcePath $reportFile -ArchivePath $archiveFile

#endregion Initialization

#region Collection Phase

$allInventoryData = [System.Collections.Generic.List[PSCustomObject]]::new()
$perVCenterData = @{}

$successCount = 0
$failCount = 0

foreach ($vc in $config.VCenters) {
    $vcName = $vc.Name
    $vcAlias = if ($vc.Alias) { $vc.Alias } else { $vcName.Split('.')[0] }
    Write-Host "`nProcessing vCenter: $vcAlias ($vcName)" -ForegroundColor Cyan

    $connection = $null
    try {
        # Retrieve credential — use parameter if provided, otherwise load from DPAPI file
        if ($VCenterCredential) {
            $credential = $VCenterCredential
            Write-Verbose "Using credential passed via -VCenterCredential parameter"
        }
        else {
            $credPath = Join-Path $credDir $vc.CredentialFile
            if (-not (Test-Path -Path $credPath)) {
                Write-Error "Credential file not found: $credPath. Run Initialize-VCenterCredentials.ps1 first."
            }
            Write-Verbose "Loading credential from '$credPath'"
            $credential = Import-Clixml -Path $credPath -ErrorAction Stop
        }

        # Connect to vCenter
        Write-Host "  Connecting to $vcName..." -ForegroundColor Gray
        $connection = Connect-VIServer -Server $vcName -Credential $credential -ErrorAction Stop
        Write-Host "  Connected to $vcName." -ForegroundColor Green

        # Collect VM inventory
        Write-Host "  Collecting VM inventory..." -ForegroundColor Gray
        $allVMs = Get-VM -Server $connection -ErrorAction Stop
        # Filter out gray-status (disconnected/orphaned) VMs
        $allVMs = @($allVMs | Where-Object { $_.ExtensionData.Summary.OverallStatus -ne 'gray' })
        Write-Verbose "  Found $($allVMs.Count) VM(s) on $vcName (after filtering gray status)"

        $inventoryData = foreach ($vm in $allVMs) {
            Write-Verbose "    Processing VM: $($vm.Name)"
            $vmView = Get-View -VIObject $vm -Property Config, Guest, Runtime, Summary

            # Template — collect only basic properties
            if ($vmView.Config.Template) {
                $tProps = [ordered]@{
                    'Name'                    = $vm.Name
                    'PowerState'              = $vmView.Summary.Runtime.PowerState
                    'Cluster'                 = ''
                    'Configured_Guest_OS'     = $vmView.Config.GuestFullName
                    'Running_Guest_OS'        = $vmView.Guest.GuestFullName
                    'Notes'                   = ''
                    'Floppydrive'             = $false
                    'USB Controller'          = $false
                    'No USB Cont'             = 0
                    'Needs Disk Consolidated' = $false
                    'NumCpu'                  = $vmView.Summary.Config.NumCpu
                    'CPU Hot Add'             = $vmView.Config.CpuHotAddEnabled
                    'Memory GB'               = [math]::Round($vmView.Summary.Config.MemorySizeMB / 1024, 2)
                    'Memory Hot Add'          = $vmView.Config.MemoryHotAddEnabled
                    'Hardware Version'        = $vmView.Config.Version
                    'VMTools Status'          = ''
                    'Datacenter'              = ''
                    'CD Drive'                = $false
                    'CD Connected'            = $false
                    'Template'                = $true
                    'IP #1'                   = ''
                    'IP #2'                   = ''
                    'IP #3'                   = ''
                    '1st vNic Type'           = ''
                    '2nd vNic Type'           = ''
                    '3rd vNic Type'           = ''
                    '4th vNic Type'           = ''
                    '5th vNic Type'           = ''
                    'VMTools Version'         = $vmView.Summary.Guest.ToolsVersionStatus
                    'Disk1'                   = ''
                    'Disk2'                   = ''
                    'Disk3'                   = ''
                    'Disk4'                   = ''
                    'Disk5'                   = ''
                    'Disk6'                   = ''
                    'Disk7'                   = ''
                    'Disk8'                   = ''
                    'Disk9'                   = ''
                    'Disk10'                  = ''
                    'Disk11'                  = ''
                    'Disk12'                  = ''
                    'vCenter'                 = $vcName
                    'Firmware'                = $vmView.Config.Firmware
                    'NestedHVEnabled'         = $false
                    'EfiSecureBootEnabled'    = $false
                    'VvtdEnabled'             = $false
                    'VbsEnabled'              = $false
                    'Guest Hostname'          = ''
                }
                # Add empty tag columns for consistent headers
                foreach ($tagDef in $config.RequiredTags) {
                    for ($t = 0; $t -lt $tagDef.Columns; $t++) {
                        $suffix = if ($tagDef.Columns -gt 1) { ($t + 1) } else { '' }
                        $tProps["$($tagDef.DisplayName)_Tag$suffix"] = ''
                    }
                }
                $tProps['Change_Number'] = ''
                $tProps['vTPM'] = $false
                $tProps['ResourcePool'] = ''
                $tProps['FolderName'] = ''
                $tProps['LastBackup'] = ''
                [PSCustomObject]$tProps
                continue
            }

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

            # VMTools version — descriptive string based on status
            $toolsVersionStatus = $vmView.Guest.ToolsVersionStatus
            $toolsVersion = switch ($toolsVersionStatus) {
                'guestToolsNotInstalled' { 'VMTools Not Installed' }
                'guestToolsUnmanaged'    { 'Guest Managed' }
                'guestToolsCurrent'      { $vmView.Guest.ToolsVersion }
                'guestToolsNeedUpgrade'  { $vmView.Guest.ToolsVersion }
                default                  { $vmView.Guest.ToolsVersion }
            }
            if ($vm.PowerState -eq 'PoweredOff' -and $configuredOS -like '*Windows*') {
                $toolsVersion = 'Windows VM Powered Off'
            }

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

            # Network adapters (up to 5) — NIC type (e.g. Vmxnet3, E1000)
            $netAdapters = Get-NetworkAdapter -VM $vm -ErrorAction SilentlyContinue | Sort-Object Name
            $nics = @('')*5
            for ($i = 0; $i -lt [math]::Min($netAdapters.Count, 5); $i++) {
                $nics[$i] = $netAdapters[$i].Type
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
            $vvtd = $vmView.Config.Flags.VvtdEnabled
            $vbs = $vmView.Config.Flags.VbsEnabled

            # Hostname vs VM name comparison — strip domain before comparing
            $guestHostname = $vmView.Guest.HostName
            if ($guestHostname) {
                $shortName = $guestHostname.Split('.')[0]
                $nameNotEqual = if ($shortName -ne $vm.Name) { $guestHostname } else { '' }
            }
            else {
                $nameNotEqual = 'NO hostname returned from vm'
            }

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

            # Helper to get tag values by category
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

            # Build base properties (ordered)
            $props = [ordered]@{
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
                'Template'                = $false
                'IP #1'                   = $ips[0]
                'IP #2'                   = $ips[1]
                'IP #3'                   = $ips[2]
                '1st vNic Type'           = $nics[0]
                '2nd vNic Type'           = $nics[1]
                '3rd vNic Type'           = $nics[2]
                '4th vNic Type'           = $nics[3]
                '5th vNic Type'           = $nics[4]
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
                'Guest Hostname'          = $nameNotEqual
            }

            # Add tag columns dynamically from config
            # vSphere category = "{TagPrefix}-{TagEnvironment}-{Category}"
            # Spreadsheet column = "{DisplayName}_Tag{N}"
            $tagEnv = $vc.TagEnvironment
            foreach ($tagDef in $config.RequiredTags) {
                $vSphereCategory = "$($config.TagPrefix)-$tagEnv-$($tagDef.Category)"
                for ($t = 0; $t -lt $tagDef.Columns; $t++) {
                    $suffix = if ($tagDef.Columns -gt 1) { ($t + 1) } else { '' }
                    $propName = "$($tagDef.DisplayName)_Tag$suffix"
                    $props[$propName] = & $getTag $vSphereCategory $t
                }
            }

            # Add remaining properties after tags
            $props['Change_Number'] = $changeNumber
            $props['vTPM'] = $vtpm
            $props['ResourcePool'] = if ($vm.ResourcePool) { $vm.ResourcePool.Name } else { '' }
            $props['FolderName'] = if ($vm.Folder) {
                $folderPath = @()
                $f = $vm.Folder
                while ($f -and $f.Name -ne 'vm') {
                    $folderPath += $f.Name
                    $f = $f.Parent
                }
                [array]::Reverse($folderPath)
                $folderPath -join '/'
            } else { '' }
            $props['LastBackup'] = $lastBackup

            [PSCustomObject]$props
        }

        # Store collected data
        $vcInventory = @($inventoryData)
        foreach ($item in $vcInventory) {
            $allInventoryData.Add($item)
        }
        $perVCenterData[$vcAlias] = $vcInventory

        Write-Host "  Collected $($vcInventory.Count) VM(s)." -ForegroundColor Green
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

#endregion Collection Phase

#region Build Workbook

if ($allInventoryData.Count -eq 0) {
    Write-Warning "No VM data collected from any vCenter. Skipping workbook creation."
    Stop-Transcript
    return
}

Write-Host "`nBuilding workbook with $($allInventoryData.Count) total VM(s)..." -ForegroundColor Cyan

# Build tag column names for reference (uses DisplayName for column headers)
$tagColumnNames = [System.Collections.Generic.List[string]]::new()
foreach ($tagDef in $config.RequiredTags) {
    for ($t = 0; $t -lt $tagDef.Columns; $t++) {
        $suffix = if ($tagDef.Columns -gt 1) { ($t + 1) } else { '' }
        $tagColumnNames.Add("$($tagDef.DisplayName)_Tag$suffix")
    }
}

# Build filtered views
$missingTagsData = foreach ($vm in $allInventoryData) {
    $missingCategories = [System.Collections.Generic.List[string]]::new()
    foreach ($tagDef in $config.RequiredTags) {
        $allEmpty = $true
        for ($t = 0; $t -lt $tagDef.Columns; $t++) {
            $suffix = if ($tagDef.Columns -gt 1) { ($t + 1) } else { '' }
            $propName = "$($tagDef.DisplayName)_Tag$suffix"
            if ($vm.$propName -ne '') {
                $allEmpty = $false
                break
            }
        }
        if ($allEmpty) {
            $missingCategories.Add($tagDef.DisplayName)
        }
    }
    if ($missingCategories.Count -gt 0) {
        $vmCopy = $vm.PSObject.Copy()
        $vmCopy | Add-Member -NotePropertyName 'MissingTagCategories' -NotePropertyValue ($missingCategories -join ', ') -Force
        $vmCopy
    }
}

$vmBiosData = @($allInventoryData | Where-Object { $_.Firmware -eq 'bios' })
$poweredOffData = @($allInventoryData | Where-Object { $_.PowerState -eq 'PoweredOff' })
$cpuHotAddFalse = @($allInventoryData | Where-Object { $_.'CPU Hot Add' -eq $false })
$memHotAddFalse = @($allInventoryData | Where-Object { $_.'Memory Hot Add' -eq $false })
$floppyDriveVMs = @($allInventoryData | Where-Object { $_.Floppydrive -eq $true })

# VMTools version — find latest numeric version and filter VMs not on it
$toolsVersionList = @($allInventoryData.'VMTools Version' | Select-Object -Unique |
    Where-Object { $_ -match '^\d[\d.]*$' })
$latestToolsVersion = ''
if ($toolsVersionList.Count -gt 0) {
    try {
        $latestToolsVersion = [string]($toolsVersionList |
            ForEach-Object { [version]$_ } | Sort-Object -Descending | Select-Object -First 1)
    }
    catch {
        $latestToolsVersion = ($toolsVersionList | Sort-Object -Descending | Select-Object -First 1)
    }
}
$vmToolsOutdated = if ($latestToolsVersion) {
    @($allInventoryData | Where-Object {
        $_.'VMTools Version' -ne $latestToolsVersion -and $_.'VMTools Version' -match '^\d[\d.]*$'
    })
} else { @() }

# Conditional formatting rules (column letters will be set based on actual data)
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

# Export All_VMs tab first (this creates the workbook)
$tempXlsx = Join-Path $OutputDir 'VMInventory_All.tmp.xlsm'
if (Test-Path $tempXlsx) { Remove-Item $tempXlsx -Force }
if (Test-Path $reportFile) { Remove-Item $reportFile -Force }

$allInventoryData | Export-Excel -Path $tempXlsx -WorksheetName 'All_VMs' `
    -AutoSize -FreezePane 2, 2 -BoldTopRow -ConditionalText $vmCfRules `
    -TableName 'All_VMs' -TableStyle Medium9

# Export additional data tabs to the same workbook
if (@($missingTagsData).Count -gt 0) {
    $missingTagsData | Export-Excel -Path $tempXlsx -WorksheetName 'MissingTags' `
        -AutoSize -FreezePane 2, 2 -BoldTopRow `
        -TableName 'MissingTags' -TableStyle Medium9
}
else {
    # Create empty sheet with header note
    Export-Excel -Path $tempXlsx -WorksheetName 'MissingTags' -InputObject $null
}

if ($vmBiosData.Count -gt 0) {
    $vmBiosData | Export-Excel -Path $tempXlsx -WorksheetName 'VM_BIOS' `
        -AutoSize -FreezePane 2, 2 -BoldTopRow -ConditionalText $vmCfRules `
        -TableName 'VM_BIOS' -TableStyle Medium9
}
else {
    Export-Excel -Path $tempXlsx -WorksheetName 'VM_BIOS' -InputObject $null
}

if ($poweredOffData.Count -gt 0) {
    $poweredOffData | Export-Excel -Path $tempXlsx -WorksheetName 'VMs_Powered_Off' `
        -AutoSize -FreezePane 2, 2 -BoldTopRow -ConditionalText $vmCfRules `
        -TableName 'VMs_Powered_Off' -TableStyle Medium9
}
else {
    Export-Excel -Path $tempXlsx -WorksheetName 'VMs_Powered_Off' -InputObject $null
}

# Per-vCenter tabs
foreach ($vcName in $perVCenterData.Keys) {
    $vcData = $perVCenterData[$vcName]
    # Sanitize vCenter name for worksheet name (max 31 chars, no special chars)
    $tabName = $vcName -replace '[:\\/\?\*\[\]]', '_'
    if ($tabName.Length -gt 31) { $tabName = $tabName.Substring(0, 31) }

    if ($vcData.Count -gt 0) {
        $vcData | Export-Excel -Path $tempXlsx -WorksheetName $tabName `
            -AutoSize -FreezePane 2, 2 -BoldTopRow -ConditionalText $vmCfRules `
            -TableName ($tabName -replace '[^A-Za-z0-9_]', '_') -TableStyle Medium9
    }
    else {
        Export-Excel -Path $tempXlsx -WorksheetName $tabName -InputObject $null
    }
}

# CPU Hot Add FALSE
if ($cpuHotAddFalse.Count -gt 0) {
    $cpuHotAddFalse | Export-Excel -Path $tempXlsx -WorksheetName 'CPU_HotAdd_FALSE' `
        -AutoSize -FreezePane 2, 2 -BoldTopRow -ConditionalText $vmCfRules `
        -TableName 'CPU_HotAdd_FALSE' -TableStyle Medium9
}
else {
    Export-Excel -Path $tempXlsx -WorksheetName 'CPU_HotAdd_FALSE' -InputObject $null
}

# Memory Hot Add FALSE
if ($memHotAddFalse.Count -gt 0) {
    $memHotAddFalse | Export-Excel -Path $tempXlsx -WorksheetName 'Memory_HotAdd_FALSE' `
        -AutoSize -FreezePane 2, 2 -BoldTopRow -ConditionalText $vmCfRules `
        -TableName 'Memory_HotAdd_FALSE' -TableStyle Medium9
}
else {
    Export-Excel -Path $tempXlsx -WorksheetName 'Memory_HotAdd_FALSE' -InputObject $null
}

# VMTools Version — VMs not on the latest tools version
if ($vmToolsOutdated.Count -gt 0) {
    $vmToolsOutdated | Export-Excel -Path $tempXlsx -WorksheetName 'VMToolsVersion' `
        -AutoSize -FreezePane 2, 2 -BoldTopRow -ConditionalText $vmCfRules `
        -TableName 'VMToolsVersion' -TableStyle Medium9
}
else {
    Export-Excel -Path $tempXlsx -WorksheetName 'VMToolsVersion' -InputObject $null
}

# Floppy Drives
if ($floppyDriveVMs.Count -gt 0) {
    $floppyDriveVMs | Export-Excel -Path $tempXlsx -WorksheetName 'FloppyDrives' `
        -AutoSize -FreezePane 2, 2 -BoldTopRow -ConditionalText $vmCfRules `
        -TableName 'FloppyDrives' -TableStyle Medium9
}
else {
    Export-Excel -Path $tempXlsx -WorksheetName 'FloppyDrives' -InputObject $null
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
$searchWs.Cells[1, 1].Value = 'VM Inventory - Search'
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
Write-Host "    All_VMs          : $($allInventoryData.Count) VM(s)" -ForegroundColor Gray
Write-Host "    MissingTags      : $(@($missingTagsData).Count) VM(s)" -ForegroundColor Gray
Write-Host "    VM_BIOS          : $($vmBiosData.Count) VM(s)" -ForegroundColor Gray
Write-Host "    VMs_Powered_Off  : $($poweredOffData.Count) VM(s)" -ForegroundColor Gray
Write-Host "    CPU_HotAdd_FALSE : $($cpuHotAddFalse.Count) VM(s)" -ForegroundColor Gray
Write-Host "    Memory_HotAdd_FALSE: $($memHotAddFalse.Count) VM(s)" -ForegroundColor Gray
Write-Host "    VMToolsVersion   : $($vmToolsOutdated.Count) VM(s) (latest: $latestToolsVersion)" -ForegroundColor Gray
Write-Host "    FloppyDrives     : $($floppyDriveVMs.Count) VM(s)" -ForegroundColor Gray
foreach ($vcName in $perVCenterData.Keys) {
    Write-Host "    $vcName : $($perVCenterData[$vcName].Count) VM(s)" -ForegroundColor Gray
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
Write-Host "End of Get-AllVMInventory.ps1 at $($endTime.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor Cyan

Stop-Transcript

#endregion Summary
