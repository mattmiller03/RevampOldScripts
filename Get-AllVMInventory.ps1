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

    $reportFile = Join-Path $OutputDir "VMInventory_$vcName.xlsx"
    $archiveFile = Join-Path $ArchiveDir "VMInventory_$vcName.xlsx"

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

        if (Test-Path $reportFile) { Remove-Item $reportFile -Force }
        $inventoryData | Export-Excel -Path $reportFile -WorksheetName 'VMInventory' `
            -AutoSize -FreezeTopRow -BoldTopRow -ConditionalText $vmCfRules
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
