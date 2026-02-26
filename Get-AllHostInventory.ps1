#Requires -Modules VMware.PowerCLI, Microsoft.PowerShell.SecretManagement, ImportExcel

<#
.SYNOPSIS
    Collects ESX host inventory from all vCenters defined in the configuration file.

.DESCRIPTION
    Iterates through each vCenter server listed in config/vcenters.json, connects using
    credentials stored in a SecretManagement vault, collects host inventory, and saves
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
    .\Get-AllHostInventory.ps1
    Runs with default paths relative to the script directory.

.EXAMPLE
    .\Get-AllHostInventory.ps1 -OutputDir "D:\Reports\HostInventory" -Verbose
    Runs with a custom output directory and verbose logging.
#>

[CmdletBinding()]
param(
    [Parameter()]
    [string]$ConfigFile = (Join-Path $PSScriptRoot 'config\vcenters.json'),

    [Parameter()]
    [string]$OutputDir = (Join-Path $PSScriptRoot 'Output\HostInventory'),

    [Parameter()]
    [string]$ArchiveDir = (Join-Path $PSScriptRoot 'Output\HostInventory\Archive'),

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
$transcriptPath = Join-Path $TranscriptDir "Get-AllHostInventory_$(Get-Date -Format 'yyyy-MM-dd_HHmmss').log"
Start-Transcript -Path $transcriptPath

$startTime = Get-Date
Write-Host "Start of Get-AllHostInventory.ps1 at $($startTime.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor Cyan

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

    $reportFile = Join-Path $OutputDir "ESX_HostInventory_$vcName.xlsx"
    $archiveFile = Join-Path $ArchiveDir "ESX_HostInventory_$vcName.xlsx"

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

        # Collect host inventory
        Write-Host "  Collecting host inventory..." -ForegroundColor Gray
        $vmHosts = Get-VMHost -Server $connection -ErrorAction Stop
        Write-Verbose "  Found $($vmHosts.Count) host(s) on $vcName"

        $vcVersion = $connection.Version
        $inventoryData = foreach ($vmHost in $vmHosts) {
            Write-Verbose "    Processing host: $($vmHost.Name)"
            $hostView = Get-View -VIObject $vmHost -Property Config, Hardware, Runtime, Summary, Network
            $services = Get-VMHostService -VMHost $vmHost -ErrorAction SilentlyContinue

            # VM counts
            $allVMs = Get-VM -Location $vmHost -ErrorAction SilentlyContinue
            $poweredOnVMs = $allVMs | Where-Object { $_.PowerState -eq 'PoweredOn' }

            # vCPU totals
            $totalVCPUs = ($allVMs | Measure-Object -Property NumCpu -Sum).Sum
            $poweredOnVCPUs = ($poweredOnVMs | Measure-Object -Property NumCpu -Sum).Sum
            $poweredOnMhz = ($poweredOnVMs | ForEach-Object { (Get-View -VIObject $_ -Property Runtime).Runtime } |
                Where-Object { $null -ne $_ } | ForEach-Object { 0 } | Measure-Object -Sum).Sum
            # Use host-level CPU usage as a practical measure for powered-on MHz
            $poweredOnMhz = $vmHost.CpuUsageMhz

            # vMemory totals
            $totalVMemGB = ($allVMs | Measure-Object -Property MemoryGB -Sum).Sum
            $poweredOnVMemGB = ($poweredOnVMs | Measure-Object -Property MemoryGB -Sum).Sum

            # Management network
            $mgmtVmk = Get-VMHostNetworkAdapter -VMHost $vmHost -VMKernel -ErrorAction SilentlyContinue |
                Where-Object { $_.ManagementTrafficEnabled } | Select-Object -First 1
            $mgmtIP = if ($mgmtVmk) { $mgmtVmk.IP } else { '' }
            $mgmtVlanId = ''
            if ($mgmtVmk) {
                # Try standard port group first, then distributed
                $pg = Get-VirtualPortGroup -VMHost $vmHost -Standard -Name $mgmtVmk.PortGroupName -ErrorAction SilentlyContinue
                if ($pg) {
                    $mgmtVlanId = $pg.VLanId
                }
                else {
                    $dpg = Get-VDPortgroup -Name $mgmtVmk.PortGroupName -ErrorAction SilentlyContinue
                    if ($dpg) { $mgmtVlanId = $dpg.VlanConfiguration.VlanId }
                }
            }

            # VMotion and FT IPs
            $vmkAdapters = Get-VMHostNetworkAdapter -VMHost $vmHost -VMKernel -ErrorAction SilentlyContinue
            $vmotionIP = ($vmkAdapters | Where-Object { $_.VMotionEnabled } | Select-Object -First 1).IP
            $ftIP = ($vmkAdapters | Where-Object { $_.FaultToleranceLoggingEnabled } | Select-Object -First 1).IP
            $vmkGateway = $hostView.Config.Network.IpRouteConfig.DefaultGateway

            # Host authentication
            $authInfo = Get-VMHostAuthentication -VMHost $vmHost -ErrorAction SilentlyContinue
            $hostAuth = if ($authInfo -and $authInfo.Domain) { $authInfo.Domain } else { 'Local' }

            # Advanced settings
            $advSettings = Get-AdvancedSetting -Entity $vmHost -ErrorAction SilentlyContinue
            $syslogServer = ($advSettings | Where-Object { $_.Name -eq 'Syslog.global.logHost' }).Value
            $hostAgentSetting = ($advSettings | Where-Object { $_.Name -eq 'Config.HostAgent.plugins.solo.enableMob' }).Value
            $ipv6Enabled = ($advSettings | Where-Object { $_.Name -eq 'Net.IPv6Enabled' }).Value

            # Services
            $sshEnabled = ($services | Where-Object { $_.Key -eq 'TSM-SSH' }).Running
            $shellEnabled = ($services | Where-Object { $_.Key -eq 'TSM' }).Running
            $dcuiEnabled = ($services | Where-Object { $_.Key -eq 'DCUI' }).Running

            # Dump collector
            $dumpCollector = ($advSettings | Where-Object { $_.Name -eq 'VMkernel.Boot.netDumpAddr' }).Value

            # Uptime
            $bootTime = $hostView.Runtime.BootTime
            $uptime = if ($bootTime) { (New-TimeSpan -Start $bootTime -End (Get-Date)).ToString('dd\.hh\:mm\:ss') } else { '' }

            # Hardware details
            $serviceTag = ''
            if ($hostView.Hardware.SystemInfo.OtherIdentifyingInfo) {
                $svcTagEntry = $hostView.Hardware.SystemInfo.OtherIdentifyingInfo |
                    Where-Object { $_.IdentifierType.Key -eq 'ServiceTag' }
                if ($svcTagEntry) { $serviceTag = $svcTagEntry.IdentifierValue }
            }

            # EVC
            $cluster = Get-Cluster -VMHost $vmHost -ErrorAction SilentlyContinue
            $maxEvcKey = if ($cluster) { $cluster.EVCMode } else { '' }

            # Physical NICs
            $pNics = (Get-VMHostNetworkAdapter -VMHost $vmHost -Physical -ErrorAction SilentlyContinue).Count

            # Hyperthreading
            $htActive = $hostView.Config.HyperThread.Active
            $logicalProcessors = $vmHost.NumCpu

            # License
            $licenseKey = ''
            $licMgr = Get-View -Id 'LicenseManager-ha-license-manager' -ErrorAction SilentlyContinue
            if ($licMgr) {
                $hostLic = $licMgr.Licenses | Where-Object {
                    $_.Properties | Where-Object { $_.Key -eq 'EntityId' -and $_.Value -eq $vmHost.MoRef.Value }
                } | Select-Object -First 1
                if ($hostLic) { $licenseKey = $hostLic.LicenseKey }
            }
            if (-not $licenseKey) {
                $licenseKey = ($advSettings | Where-Object { $_.Name -eq 'License.ProductKey' }).Value
            }

            # Custom attributes
            $customAttribs = Get-Annotation -Entity $vmHost -ErrorAction SilentlyContinue
            $ebsNumber = ($customAttribs | Where-Object { $_.Name -eq 'EBS_Number' }).Value
            $dlaAsset = ($customAttribs | Where-Object { $_.Name -eq 'DLA_Asset' }).Value
            $siteLocation = ($customAttribs | Where-Object { $_.Name -eq 'Site_Location' }).Value

            # Secure boot and TPM
            $secureBoot = $false
            $tpmSupport = $false
            $tpmVersion = ''
            if ($hostView.Runtime.BootInfo) {
                $secureBoot = [bool]$hostView.Runtime.BootInfo.SecureBoot
            }
            if ($hostView.Hardware.TpmInfo) {
                $tpmSupport = $true
                $tpmVersion = $hostView.Hardware.TpmInfo.TpmVersion
            }

            [PSCustomObject]@{
                'Name'                       = $vmHost.Name
                'ESXi-Version'               = $vmHost.Version
                'Build-Version'              = $vmHost.Build
                'Management IP'              = $mgmtIP
                'vLan ID'                    = $mgmtVlanId
                'PowerState'                 = $vmHost.PowerState
                'Manufacturer'               = $vmHost.Manufacturer
                'Model'                      = $vmHost.Model
                'Service_Tag'                = $serviceTag
                'Total_VMs'                  = @($allVMs).Count
                'PoweredOnVMss'              = @($poweredOnVMs).Count
                'ProcessorType'              = $vmHost.ProcessorType
                'CPU_Sockets'                = $hostView.Hardware.CpuInfo.NumCpuPackages
                'Cores_per_Socket'           = $hostView.Hardware.CpuInfo.NumCpuCores / $hostView.Hardware.CpuInfo.NumCpuPackages
                'CPU_Cores'                  = $hostView.Hardware.CpuInfo.NumCpuCores
                'TotalHost_Mhz'              = $vmHost.CpuTotalMhz
                'AssignedTotal_vCPUs'        = $totalVCPUs
                'PoweredOn_vCPUs'            = $poweredOnVCPUs
                'PoweredOn_Mhz'              = $poweredOnMhz
                'Memory(GB)'                 = [math]::Round($vmHost.MemoryTotalGB, 2)
                'AssignedTotal-vMemory(GB)'  = [math]::Round($totalVMemGB, 2)
                'PoweredOn-vMemory(GB)'      = [math]::Round($poweredOnVMemGB, 2)
                'Host Authentication'        = $hostAuth
                'Max-EVC-Key'                = $maxEvcKey
                'Cluster'                    = if ($cluster) { $cluster.Name } else { '' }
                'DataCenter'                 = (Get-Datacenter -VMHost $vmHost -ErrorAction SilentlyContinue).Name
                'vCenter Server'             = $vcName
                'vCenter Version'            = $vcVersion
                'Esxi-status'                = $vmHost.ConnectionState
                'Physical-NICs'              = $pNics
                'ESXi Shell-Enabled'         = $shellEnabled
                'SSH-Enabled'                = $sshEnabled
                'DCUI-Enabled'               = $dcuiEnabled
                'Uptime'                     = $uptime
                'Syslog-Server'              = $syslogServer
                'Dump-Collector'             = $dumpCollector
                'Config.HostAgent Setting'   = $hostAgentSetting
                'Hyperthread Active'         = $htActive
                'Logical Processors'         = $logicalProcessors
                'VMotion IP'                 = $vmotionIP
                'Fault Tolerance IP'         = $ftIP
                'License Key'                = $licenseKey
                'ConnectionState'            = $vmHost.ConnectionState
                'vmKernel Gateway'           = $vmkGateway
                'EBS_Number'                 = $ebsNumber
                'DLA_Asset'                  = $dlaAsset
                'Site_Location'              = $siteLocation
                'IPv6 Enabled'               = $ipv6Enabled
                'SecureBoot'                 = $secureBoot
                'TPMSupport'                 = $tpmSupport
                'TPMVersion'                 = $tpmVersion
            }
        }

        # Conditional formatting rules
        $hostCfRules = @(
            # Red: SSH enabled
            New-ConditionalText -Text 'True' -Range 'AF:AF' -BackgroundColor Red -ConditionalTextColor White
            # Red: ESXi Shell enabled
            New-ConditionalText -Text 'True' -Range 'AE:AE' -BackgroundColor Red -ConditionalTextColor White
            # Red: Disconnected or NotResponding
            New-ConditionalText -Text 'Disconnected' -Range 'AC:AC' -BackgroundColor Red -ConditionalTextColor White
            New-ConditionalText -Text 'NotResponding' -Range 'AC:AC' -BackgroundColor Red -ConditionalTextColor White
            # Red: ConnectionState issues
            New-ConditionalText -Text 'Disconnected' -Range 'AQ:AQ' -BackgroundColor Red -ConditionalTextColor White
            New-ConditionalText -Text 'NotResponding' -Range 'AQ:AQ' -BackgroundColor Red -ConditionalTextColor White
            # Yellow: PowerState not PoweredOn
            New-ConditionalText -Text 'Standby' -Range 'F:F' -BackgroundColor Yellow
            New-ConditionalText -Text 'PoweredOff' -Range 'F:F' -BackgroundColor Yellow
            # Yellow: Maintenance mode
            New-ConditionalText -Text 'Maintenance' -Range 'AC:AC' -BackgroundColor Yellow
        )

        $inventoryData | Export-Excel -Path $reportFile -WorksheetName 'HostInventory' `
            -AutoSize -FreezeTopRow -BoldTopRow -ConditionalText $hostCfRules -Force
        Write-Host "  Collected $(@($inventoryData).Count) host(s)." -ForegroundColor Green

        Write-Host "  Report saved: $reportFile" -ForegroundColor Green
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
Write-Host "End of Get-AllHostInventory.ps1 at $($endTime.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor Cyan

Stop-Transcript

#endregion Summary
