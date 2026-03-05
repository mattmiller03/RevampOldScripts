#PURPOSE: This script will collect all ESX inventory from all vCenter servers

#CHANGELOG
#Version 1.00 - 07/26/24 - MDR - Initial version
#Version 1.01 - 08/08/24 - MDR - Add parameter for passing the vCenter password
#Version 1.02 - 08/15/24 - MDR - If $ReportPath gets an access denied then exit the script
#Version 1.03 - 09/12/24 - MDR - Passing svc_vrapsh credential for all interactions with the filer
#Version 1.04 - 09/12/24 - MDR - Modify how the script runs depending on if run manually or from Orchestrator
#Version 1.05 - 09/16/24 - MDR - Renaming the output file, changing report path directory, moving older versions of this report to SavedFiles folder, renaming $vCenterCreds variable to $vCenterPassword
#Version 1.06 - 10/03/24 - MDR - Importing Base64 passwords from .csv rather than passing as parameters to improve security
#Version 1.07 - 10/04/23 - MDR - Minor Corrections
#Version 1.08 - 10/25/24 - MDR - Added "-FreezePane 2, 2" to all Export-Excel lines
#Version 1.09 - 05/01/25 - MDR - Updated with the new Dev-Test vCenter server
#Version 1.10 - 10/08/25 - MDR - Store report to local temp folder before moving it to the network share

Param ( $ParameterPath, $RunFromOrchestrator )

#Version 1.07 - Only import parameters from file if run from Orchestrator
If ($RunFromOrchestrator -eq "True") {
    #Version 1.06 - Import Base64 passwords from CSV
    $Base64Passwords = Import-CSV $ParameterPath

    #Version 1.06 - Delete the CSV file since it is no longer needed and for security reasons
    Remove-Item $ParameterPath | Out-Null

    #Version 1.06 - Store passwords to temp variables
    $vCenterPasswordBase64 = $Base64Passwords.vCenterPassword
    $VRAPasswordBase64 = $Base64Passwords.VRAPassword

    #Version 1.06 - Decode passwords from Base64 to plain text
    $vCenterPassword = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($vCenterPasswordBase64))
    $VRAPassword = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($VRAPasswordBase64))
}

#Configure variables
$TodaysDate = Get-Date -Format "MMddyyyy" #Version 1.05 - Made this a 4 digit year
$ReportPath = "\\orgaze.dir.ad.dla.mil\DCC\VirtualTeam\Reports" #Version 1.05 - Changed path to VPM share \ Reports folder
$ReportFileName = "Master_ESX_HostInventory_$TodaysDate.xlsx" #Version 1.05 - Storing report file name in a variable instead of hardcoding throughout the script
$ShortReportName = "Master_ESX_HostInventory" #Version 1.05 - Adding a short name which is used when finding older versions of this report to move to SavedFiles
$ESXInventoryData = New-Object System.Collections.Generic.List[System.Object]

$ReportColumnHeaders = "Name", "ESXi-Version", "Build-Version", "Management IP", "vLan ID", "PowerState", "Manufacturer", "Model", "Service_Tag", "Total_VMs", "PoweredOnVMss", "ProcessorType", "CPU_Sockets",
                    "Cores_per_Socket", "CPU_Cores", "TotalHost_Mhz", "AssignedTotal_vCPUs", "PoweredOn_vCPUs", "PoweredOn_Mhz", "Memory(GB)", "AssignedTotal-vMemory(GB)", "PoweredOn-vMemory(GB)",
                    "Host Authentication", "Max-EVC-Key", "Cluster", "DataCenter", "vCenter Server", "vCenter Version", "Esxi-status", "Physical-NICs", "ESXi Shell-Enabled", "SSH-Enabled", "DCUI-Enabled",
                    "Uptime", "Syslog-Server", "Dump-Collector", "Config.HostAgent Setting", "Hyperthread Active", "Logical Processors", "VMotion IP", "Fault Tolerance IP", "License Key", "ConnectionState",
                    "vmKernel Gateway", "EBS_Number", "DLA_Asset", "Site_Location", "IPv6 Enabled", "SecureBoot", "TPMSupport", "TPMVersion"

#Version 1.10 - Ensure Temp folder exists
If (!(Test-Path "C:\Temp")) {
    New-Item "C:\Temp" -ItemType Directory | Out-Null
}

#Version 1.01 - Create an ouput file for commands being executed
Start-Transcript C:\Temp\Get_All_ESX_Inventory_$TodaysDate.txt

Clear-Host

#If a ReportPath drive mapping already exists then remove it
Get-PSDrive -Name ReportPath -ErrorAction SilentlyContinue | Remove-PSDrive

#Version 1.04 - If this is run from Orchestrator then generate a credential to perform the PSDrive mapping
If ($RunFromOrchestrator -eq "True") {
    #Version 1.03 - Create a credential for svc_vrapsh
    $VRACred = New-Object System.Management.Automation.PSCredential -ArgumentList @("DIR\svc_vrapsh",(ConvertTo-SecureString -String $VRAPassword -AsPlainText -Force))
    #Version 1.03 - Map a drive to the ReportPath
    New-PSDrive -Name "ReportPath" -PSProvider FileSystem -Root $ReportPath -Credential $VRACred | Out-Null
} Else { #If not run from Orchestrator then just map the PSDrive with current credentials
    #Version 1.07 - Added a silentlycontinue to prevent reporting when the drive is already mapped
    New-PSDrive -Name "ReportPath" -PSProvider FileSystem -Root $ReportPath -ErrorAction SilentlyContinue | Out-Null
}

If (!(Get-PSDrive -Name "ReportPath")) {
    Write-Host "Failed to connect to ReportPath.  Exiting" -ForegroundColor Red
    Break
}

#Version 1.02 - If for some reason an Access Denied error happens here then exit the script
Try {
    #If this report exists already then delete it
    If (Test-Path "ReportPath:\$ReportFileName" -ErrorAction Stop) {
        Remove-Item "ReportPath:\$ReportFileName"
    }
} Catch {
    Write-Host "Unable to access $ReportPath" -ForegroundColor Red
    Break
}

#Check for PowerShell 7.x.  If running an older version then exit the script
If ($PSVersionTable.PSVersion -lt [Version]"7.0.0") {
    Write-Host "You are running PowerShell version $($PSVersionTable.PSVersion) and at least 7.x is required"
    Break
}

#Check to see if ImportExcel is installed
$CheckForImportExcel = Get-Command Import-Excel -ErrorAction SilentlyContinue

#If ImportExcel is not found then prompt for folder where it is located
If (!$CheckForImportExcel) {
    Write-Host "The ImportExcel module is required for this script to run" -ForegroundColor Red
    Write-Host "`nA copy of this module is located in \\orgaze.dir.ad.dla.mil\J6_INFO_OPS\J64\J64C\WinAdmin\VulnMgt\Software\ImportExcel"
    Write-Host "`nPlace a copy of this module in C:\Program Files\WindowsPowerShell\Modules"
    Break
}

#Version 1.01 - If the vCenter password is passed then no need to prompt for it
#Version 1.05 - Changed variable from $vCenterCreds to $vCenterPassword
If ($vCenterPassword -eq $null -or $vCenterPassword -eq "") {
    #Prompt for vCenter Password
    $vCenterPassword = Read-Host "Input the vCenter password" -MaskInput
}

#Exit if no password was entered
If ($vCenterPassword -eq $null) {
    Write-Host "`nNo password was entered so exiting the script" -ForegroundColor Red
    Break
}

#Import list of all vCenter servers and additional info about those servers
$vCenterServerList = Import-CSV "\\orgaze.dir.ad.dla.mil\DCC\VirtualTeam\Scripts\MikeR\vCenter_Servers.csv"

#Loop through each vCenter server
ForEach ($vCenterServer in $vCenterServerList) {
    #Output vCenter Server name
    Write-Host "`nScanning $($vCenterServer.ServerName)"

    #Generate the vCenter credentials to connect
    $SecurePassword = ConvertTo-SecureString $vCenterPassword -AsPlainText -Force
    $vCenterSecureCreds = New-Object System.Management.Automation.PSCredential -ArgumentList ( $vCenterServer.User, $SecurePassword )

    #Disconnected from all vCenter Servers
    Try {
        Disconnect-VIServer -Server * -Confirm:$false -ErrorAction SilentlyContinue
    } Catch {}

    #Connect to vCenter Server
    Connect-VIServer $vCenterServer.ServerName -Credential $vCenterSecureCreds | Out-Null
    
    #Confirm that the connection worked
    If (!($global:DefaultVIServers | Where { $_.Name -eq $vCenterServer.ServerName })) {
        #Disconnect from all vCenter Servers
        Write-Host "`nFailed to connect to vCenter Server $($vCenterServer.ServerName)" -ForegroundColor Red
        Break
    }
    
    #Attempt to get the License assigned to each host, need a few extra variables defined here
    $ServiceInstance = Get-View ServiceInstance
    $LicenseManager = Get-View $ServiceInstance.Content.LicenseManager
    $LicenseAssignmentManager = Get-View $LicenseManager.LicenseAssignmentManager

    #Get a list of all ESX hosts
    $ESXList = Get-VMHost | Sort Name

    #Populate variables for progress bar
    $TotalESXs = $ESXList.Count
    $CurrentESXNum = 1
 
   ForEach ($ESXHost in $ESXList) {
        #Version 1.04 - If this is run from Orchestrator then output the ESX host name to get written to the transcript file
        If ($RunFromOrchestrator -eq "True") {
            Write-Host $ESXHost.Name
        } Else { #If not from Orchestrator then write a progress bar
            #Update progress bar
            Write-Progress -Activity "Getting ESX data" -Status "$CurrentESXNum of $TotalESXs" -PercentComplete ($CurrentESXNum / $TotalESXs * 100)
        }

        #EsxiHost CPU info
        $HostCPU = $ESXHost.ExtensionData.Summary.Hardware.NumCpuPkgs  
        $HostCPUcore = $ESXHost.ExtensionData.Summary.Hardware.NumCpuCores/$HostCPU
        $LogicalCPU = $ESXHost.ExtensionData.hardware.cpuinfo.NumCpuThreads
        $Hyperthread = ($ESXHost).HyperthreadingActive
        $ConnectionState = $ESXHost.ConnectionState

        #All Virtual Machines Info  
        $VMsOnHost = $ESXHost | Get-VM   
        $PoweredOnVMs = $VMsOnHost | Where-Object {$_.PowerState -eq "PoweredOn"}  
   
        #EsxiHost and VM -- CPU calculation  
        $AssignedTotalvCPU = $VMsOnHost | Measure-Object NumCpu -Sum | Select-Object -ExpandProperty sum
        $PoweredOnvCPU = $PoweredOnVMs | Measure-Object NumCpu -Sum | Select-Object -ExpandProperty sum
        $MhzPerCore = $ESXHost.CPUTotalMhz / $ESXHost.NumCpu
        $TotalPoweredOnMhz = $MhzPerCore * $PoweredOnvCPU
       
        #EsxiHost and VM -- Memory calculation  
        $TotalMemory = [math]::round($ESXHost.MemoryTotalGB)

        # Trying to eliminate errors coming from the next few lines if there are NO vm's on the host
        If ($PoweredOnVMs.Count -gt 0 ) {  
            $Calulatedvmmemory = $VMsOnHost | Measure-Object MemoryGB -sum | Select-Object -ExpandProperty sum  
            $TotalvmMemory = [math]::round($Calulatedvmmemory)  
            $Calulatedvmmemory = $PoweredOnVMs | Measure-Object MemoryGB -sum | Select-Object -ExpandProperty sum  
            $PoweredOn_vMemory = "{0:N2}" -f $Calulatedvmmemory  
        } Else { #No powered on VM on the HOST
	        $Calulatedvmmemory = 0
	        $TotalvmMemory = 0
	        $PoweredOn_vMemory = 0
        } #End of IF/ELSE Powered On VMs on the host

        #Cluster and Datastore info
        $ClusterInfo = $ESXHost | Get-Cluster
        $ClusterName = $ClusterInfo.Name
        $DataCenterInfo = Get-DataCenter -VMHost $ESXHost.Name
        $DatacenterName = $DataCenterInfo.Name

        #vCenterinfo
        $vCenter = $vCenterServer.ServerName
        $vCenterVersion = $global:DefaultVIServers | where {$_.Name -eq $vCenter} | %{"$($_.Version) build $($_.Build)"}
        $UPtime = (Get-Date) - ($ESXHost.ExtensionData.Runtime.BootTime) | Select-Object -ExpandProperty days
        $HostView = $ESXHost | Get-View
	    $ESXHostid = $HostView.config.host.value
	    $ESXLicense = ($LicenseAssignmentManager.QueryAssignedLicenses($ESXHostid))[0]
	    $ESXLicenseKey = $ESXLicense.AssignedLicense.LicenseKey
        $ConnectionState = $ESXHost.ConnectionState
        $DefaultGateway = $ESXHost.ExtensionData.Config.Network.IPRouteConfig.DefaultGateway

        #Resetting some values that are left over from the previous host inventory, until the host is Connected or in Maintenance these settings are not available
        $ManagementIP = "Not available"
        $vLanID = "Not available"
        $ServiceTag = "Not available"
        $Domain = "Not available"
        $syslog = "Not available"
        $Admins = "Not available"
 
        #NOTE: Need to put in a test for host that are not in Maintenance or Connected, and bypass those hosts
        If ($ConnectionState -eq "Connected" -or $ConnectionState -eq "Maintenance") {
            #Version 1.00 - Moved these so only connected hosts would have this info pulled, not non-responsive hosts
            $DCUIservice = $ESXHost | Get-VMHostService | Where-object {$_.key -eq "DCUI"} | Select-Object -ExpandProperty running
            $SSHservice = $ESXHost | Get-VMHostService | Where-object {$_.key -like "*ssh*"} | Select-Object -ExpandProperty running
            $ESXiservice = $ESXHost | Get-VMHostService | Where-object {$_.key -eq "Tsm"} | Select-Object -ExpandProperty running

            $ESXCLI = $ESXHost | Get-EsxCli -V2
            $ServiceTag = $ESXCLI.hardware.platform.get.Invoke().SerialNumber
            $SecureBoot = $ESXCLI.system.settings.encryption.get.Invoke().RequireSecureBoot
        
            #Esxihost Management IP and vlan ID 
            $Managementinfo = $ESXHost | Get-VMHostNetworkAdapter | Where-Object {$_.ManagementTrafficEnabled -eq $true}
            $VirtualPortGroup = $ESXHost | Get-VirtualPortGroup 
            $IPinfo = $Managementinfo | select-object -ExpandProperty ip
            $ManagementPortGroup = $Managementinfo.extensiondata.spec
            $ManagementIP = $IPinfo -join ", "  
            $DefaultGateway = $ESXHost.ExtensionData.Config.Network.IPRouteConfig.DefaultGateway
            
            #Added the following to gather the IP with VMotion and another line with Fault Tolerance Enabled
	        $VMotionInfo = $ESXHost | Get-VMHostNetworkAdapter | Where-Object {$_.VMotionEnabled -eq $true}
	        $IPinfo1 = $VMotionInfo | select-object -ExpandProperty ip
	        $VMotionIP = $IPinfo1 -join ", "
	        $FaultToleranceInfo = $ESXHost | Get-VMHostNetworkAdapter | Where-Object {$_.FaultToleranceLoggingEnabled -eq $true}
	        $IPinfo2 = $FaultToleranceInfo |select-object -ExpandProperty ip
	        $FaultToleranceIP = $IPinfo2 -join ", "
            $MulitvLans = @()

            #If there is a Mgmt Port Group
            If ($ManagementPortGroup.DistributedVirtualPort -ne $null) {  
                $vLanIDinfo = $VirtualPortGroup | Where-Object {$Managementinfo.PortGroupName -contains $_.name}  
                ForEach ($MGMTVlan in $vLanIDinfo) {  
                    $MulitvLans += $MGMTVlan.ExtensionData.config.DefaultPortConfig.Vlan.VlanId  
                } #End of foreach mgmt VLan  
                $vLanID = $MulitvLans -join ", "  
            } Else {  # If Mgmt PortGroup = NULL
                $vLanIDinfo = $VirtualPortGroup | Where-Object {$ManagementPortGroup.Portgroup -contains $_.name } | Select-Object -ExpandProperty VLanId  
                ForEach ($MGMTVlan in $vLanIDinfo) {  
                    $MulitvLans += $MGMTVlan
                }  #End of foreach mgmt VLan  
                $vLanID = $MulitvLans -join ", "  
            } #End of if/else Mgmt PortGroup

            #EsxiHost Details  
            $Domain = ($ESXHost | Get-VMHostAuthentication).Domain 
            $Admins = Get-AdvancedSetting -Entity $ESXHost | where Name -like "*esxAdminsGroup" | Select -ExpandProperty Value
            $DumpCollector = $ESXCLI.system.coredump.network.get.Invoke().NetworkServerIP 
            $syslog = ($ESXHost | Get-AdvancedSetting -Name Syslog.global.logHost).value
            #Add the newly defined Host Custom Attributes
            $ebsnum = (Get-Annotation -Entity $ESXHost -CustomAttribute EBS_Number).Value
            $assetnum = (Get-Annotation -Entity $ESXHost -CustomAttribute DLA_Asset_Number).Value
            $ipv6status = ($ESXHost | select Name, @{N='IPv6 Enabled';E={$_.ExtensionData.Config.Network.ipv6enabled}})."IPV6 Enabled"

            #05022023 Need to use the connected vCenter to generate the correct tag names, each vC different
            switch ($vCenter) {
                daisv0tp231.dir.ad.dla.mil { #Dayton DEV new tag names working, others need setup and tested
                    $LocationTagName = "vCenter-Dev-Site_Location"
                } #End of Dayton Dev vCenter
                daisv0pp241.dir.ad.dla.mil {
                    $LocationTagName = "vCenter-Prod-Site_Location"
                } #End of Dayton Prod vCenter
                trisv0pp241.dir.ad.dla.mil {
                    $LocationTagName = "vCenter-Prod-Site_Location"
                } #End of Tracy Prod vCenter
                daisv0pp271.ics.dla.mil {
                    $LocationTagName = "vCenter-OT-Site_Location"
                } #End of Dayton OT vCenter
                trisv0pp271.ics.dla.mil {
                    $LocationTagName = "vCenter-OT-Site_Location"
                } #End of Tracy OT vCenter
                daisv0pp261.dir.ad.dla.mil {
                    $LocationTagName = "vCenter-VDI-Site_Location"
                } #End of Dayton VDI 1 vCenter
                daisv0pp262.dir.ad.dla.mil {
                    $LocationTagName = "vCenter-VDI-Site_Location"
                } #End of Dayton VDI 2 vCenter
                klisv0pp251.dir.ad.dla.mil {
                    $LocationTagName = "vCenter-Kleber-Site_Location"
                } #End of Kleber Prod vCenter
                default { 
                    Write-Host "$vCenter doesn't match anything in the script" -ForegroundColor Red
                }
            } #End of switch vCenter

            #Process the TAGS on hosts
            $LocationTag = (Get-TagAssignment -Entity $ESXHost -Category $LocationTagName).Tag.Name

            #Increment the current ESX host number for the progress bar
            $CurrentESXNum++
        }
        
        $ESXInventoryData.add((New-Object "psobject" -Property @{"Name"=$ESXHost.name;"ESXi-Version"=$ESXHost.version;"Build-Version"=$ESXHost.build;"Management IP"=$ManagementIP;"vLan ID"=$vlanID;
        "PowerState"=$ESXHost.PowerState;"Manufacturer"=$ESXHost.Manufacturer;"Model"=$ESXHost.Model;"Service_Tag"=$ServiceTag;"Total_VMs"=$VMsOnHost.count;"PoweredOnVMss"=$PoweredOnVMs.Count;
        "ProcessorType"=$ESXHost.ProcessorType;"CPU_Sockets"=$HostCPU;"Cores_per_Socket"=$HostCPUcore;"CPU_Cores"=$ESXHost.Numcpu;"TotalHost_Mhz"=$ESXHost.CPUTotalMhz;"AssignedTotal_vCPUs"=$AssignedTotalvCPU;
        "PoweredOn_vCPUs"=$PoweredOnvCPU;"PoweredOn_Mhz"=$TotalPoweredOnMhz;"Memory(GB)"=$TotalMemory;"AssignedTotal-vMemory(GB)"=$TotalvmMemory;"PoweredOn-vMemory(GB)"=$PoweredOn_vMemory;
        "Host Authentication"=$Domain;"Max-EVC-Key"=$ESXHost.ExtensionData.Summary.MaxEVCModeKey;"Cluster"=$ClusterName;"DataCenter"=$DatacenterName;"vCenter Server"=$vcenter;
        "vCenter Version"=$vCenterVersion;"Esxi-status"=$ESXHost.ExtensionData.Summary.OverallStatus;"Physical-NICs"=$ESXHost.ExtensionData.summary.hardware.NumNics;"ESXi Shell-Enabled"=$ESXiservice;
        "SSH-Enabled"=$SSHservice;"DCUI-Enabled"=$DCUIservice;"Uptime"=$UPtime;"Syslog-Server"=$syslog;"Dump-Collector"=$DumpCollector;"Config.HostAgent Setting"=$Admins;"Hyperthread Active"=$Hyperthread;
        "Logical Processors"=$LogicalCPU;"VMotion IP"=$VMotionIP;"Fault Tolerance IP"=$FaultToleranceIP;"License Key"=$ESXLicenseKey;"ConnectionState"=$ConnectionState;"vmKernel Gateway"=$DefaultGateway;
        "EBS_Number"=$ebsnum;"DLA_Asset"=$assetnum;"Site_Location"=$LocationTag;"IPv6 Enabled"=$ipv6status;"SecureBoot"=$SecureBoot;"TPMSupport"=$ESXHost.ExtensionData.Capability.TpmSupported;
        "TPMVersion"=$ESXHost.ExtensionData.Capability.TpmVersion}))
    }
    
    #Close the progress bar
    Write-Progress -Activity "Getting ESX data" -Completed
}

Write-Host "`nWriting data to Excel" -ForegroundColor Cyan

#Version 1.05 - Move the old reports into the SavedFiles path
Get-ChildItem -Path $ReportPath -Filter "$ShortReportName*" | ForEach {
    Move-Item -Path $_.FullName "$ReportPath\SavedFiles" -Force
}

#Set the data headers in the correct order
$FinalOutput = $ESXInventoryData | Select $ReportColumnHeaders | Sort Name

#All ESX Hosts
$FinalOutput | Export-Excel -WorksheetName "All_Host" -Path "C:\Temp\$ReportFileName" -AutoSize -Append -FreezePane 2, 2
#Missing SecureBoot
$FinalOutput | Where { $_.SecureBoot -eq $false } | Export-Excel -WorksheetName "NOT_SecureBoot" -Path "C:\Temp\$ReportFileName" -AutoSize -Append -FreezePane 2, 2
#Not Patched - NOTE: This line needs to be manually updated as versions change
$FinalOutput | Where { $_."ESXi-Version" -ne "7.0.3" -or $_."Build-Version" -ne "23794027" } | Export-Excel -WorksheetName "Not_Patched" -Path "C:\Temp\$ReportFileName" -AutoSize -Append -FreezePane 2, 2
#Not connected
$FinalOutput | Where { $_.ConnectionState -ne "Connected" } | Export-Excel -WorksheetName "Not_Connected" -Path "C:\Temp\$ReportFileName" -AutoSize -Append -FreezePane 2, 2
#Not ESX 7
$FinalOutput | Where { $_."ESXi-Version" -notlike "7*" -and $_.ConnectionState -ne "NotResponding" } | Export-Excel -WorksheetName "Not_On_ESX_7" -Path "C:\Temp\$ReportFileName" -AutoSize -Append -FreezePane 2, 2

#Create a hash array list for vCenter Servers
$vCenterServerList = @("daisv0pp261.dir.ad.dla.mil","daisv0pp262.dir.ad.dla.mil","daisv0tp231.dir.ad.dla.mil","daisv0pp241.dir.ad.dla.mil","daisv0pp271.ics.dla.mil","klisv0pp251.dir.ad.dla.mil","trisv0pp241.dir.ad.dla.mil","trisv0pp271.ics.dla.mil")
$ExcelTabNameList = @("daisv0pp261","daisv0pp262","DaytonDev","DaytonProd","DaytonICS","Kleber","TracyProd","TracyICS")

#Loop through each vCenter Server and create a tab for them
For ($CurrvCenter = 0; $CurrvCenter -lt $vCenterServerList.Count; $CurrvCenter++) {
    $FinalOutput | Where { $_."vCenter Server" -eq $vCenterServerList[$CurrvCenter] } | Export-Excel -WorksheetName $ExcelTabNameList[$CurrvCenter] -Path "C:\Temp\$ReportFileName" -AutoSize -Append -FreezePane 2, 2
}

#Version 1.10 - Only do auto filter if not running from the PSH host
If ($RunFromOrchestrator -ne "True") {
    #Set auto filter on all worksheets
    #Open the spreadsheet that was created
    $ExcelPkg = Open-ExcelPackage -Path "C:\Temp\$ReportFileName"
    #Loop through each worksheet
    ForEach ($WorkSheet in $ExcelPkg.Workbook.Worksheets) {
        #Get the range of data in the worksheet
        $UsedRange = $WorkSheet.Dimension.Address

        #If the worksheet isn't blank
        If ($UsedRange -ne $null) {
            #Enable auto filter on that range
            $WorkSheet.Cells[$UsedRange].AutoFilter = $true
        }
    }
    #Close the spreadsheet
    Close-ExcelPackage $ExcelPkg
}

#Version 1.10 - Move the report to the report folder
Move-Item -Path "C:\Temp\$ReportFileName" -Destination "ReportPath:\"

Write-Host "`nScript Complete.  Report written to $ReportPath\ESXinventory-$TodaysDate.xlsx" -ForegroundColor Green

#Version 1.01 - Stop the transcript
Stop-Transcript
# SIG # Begin signature block
# MIIL6gYJKoZIhvcNAQcCoIIL2zCCC9cCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBRWys4j9RtX5g6
# S7yw/oYqu5m20gxqPg9snu25K2hh8qCCCS0wggRsMIIDVKADAgECAgMSNG8wDQYJ
# KoZIhvcNAQELBQAwWjELMAkGA1UEBhMCVVMxGDAWBgNVBAoTD1UuUy4gR292ZXJu
# bWVudDEMMAoGA1UECxMDRG9EMQwwCgYDVQQLEwNQS0kxFTATBgNVBAMTDERPRCBJ
# RCBDQS02MzAeFw0yMzA0MTAwMDAwMDBaFw0yNzA0MDcxMzU1NTRaMGYxCzAJBgNV
# BAYTAlVTMRgwFgYDVQQKEw9VLlMuIEdvdmVybm1lbnQxDDAKBgNVBAsTA0RvRDEM
# MAoGA1UECxMDUEtJMQwwCgYDVQQLEwNETEExEzARBgNVBAMTCkNTLkRMQS4wMDUw
# ggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCMxpcnbJBwwjRmBbyprfGQ
# i2nwvtW0H0UO192JBLpyEBkL5XxFA7iJYxXzv5sJ0whsfH8CBN6ly47Bf+QB/EpW
# Fvaay/eYc/O7uGEmk1NX2wYVH1JzrZ7hrHDSL87vcH9mwKVkbRRGVgUNvdfqtXVC
# TbsFRd1f77rzdeCTkKymd2I3Vlt+Nnr0gOy91rn+BXrDJERxeWvmeURfCfxK4D4n
# bGUW2RdsM14sejhnzw2VjrWLXezsLvWCK5rXYoTDLWD2xqrDaYqXB8V8viqKnJFW
# tw8k29z+VOH7BWrk/hZoEDTqIoLfcged0V3Vw2ivSyOnbO+JkFhZywVieBMXApYJ
# AgMBAAGjggEtMIIBKTAfBgNVHSMEGDAWgBQX5kvIGkvJp6ZwtExNXsj2NtQwmDA3
# BgNVHR8EMDAuMCygKqAohiZodHRwOi8vY3JsLmRpc2EubWlsL2NybC9ET0RJRENB
# XzYzLmNybDAOBgNVHQ8BAf8EBAMCB4AwFgYDVR0gBA8wDTALBglghkgBZQIBCyow
# HQYDVR0OBBYEFPgBvFMbp0POnSIbgh8iW8ENigzdMGUGCCsGAQUFBwEBBFkwVzAz
# BggrBgEFBQcwAoYnaHR0cDovL2NybC5kaXNhLm1pbC9zaWduL0RPRElEQ0FfNjMu
# Y2VyMCAGCCsGAQUFBzABhhRodHRwOi8vb2NzcC5kaXNhLm1pbDAfBgNVHSUEGDAW
# BgorBgEEAYI3CgMNBggrBgEFBQcDAzANBgkqhkiG9w0BAQsFAAOCAQEAClCkI904
# YRZn8KpSbGvsf8mSPsIAtHc4DrJv+8Q7a/ZCmUUjIGJMVGgWzUbik63meMbMTxG2
# RfI7c9EPb1EoowEzAnBC1ctf28PRhV//Dlaq2PeWm0gu0ozl6XD6N6GGfgqDKdwy
# 2nbInDNOjJFqgV2jeD9Pl11Ji2zTeLhc67EQWeUlb+GjOgwVooViK0Xkow/C+eQs
# DKfOZkt2HDXumJSijZ+0+GHSLrJlbAI5vB962LnKo3JTKh/VfMP/j6HfzT5nJ7rw
# 95d0s1L/Ah0B4pUiYrFkHyzX6qoMCfLh2iCPQVTg+B26dufCAAJVNOZWzBdQiVk4
# fqtL8riJSQt0tjCCBLkwggOhoAMCAQICAgUPMA0GCSqGSIb3DQEBCwUAMFsxCzAJ
# BgNVBAYTAlVTMRgwFgYDVQQKEw9VLlMuIEdvdmVybm1lbnQxDDAKBgNVBAsTA0Rv
# RDEMMAoGA1UECxMDUEtJMRYwFAYDVQQDEw1Eb0QgUm9vdCBDQSAzMB4XDTIxMDQw
# NjEzNTU1NFoXDTI3MDQwNzEzNTU1NFowWjELMAkGA1UEBhMCVVMxGDAWBgNVBAoT
# D1UuUy4gR292ZXJubWVudDEMMAoGA1UECxMDRG9EMQwwCgYDVQQLEwNQS0kxFTAT
# BgNVBAMTDERPRCBJRCBDQS02MzCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoC
# ggEBAMUSXdaAfqLX+7VK7zuVwdeUHt06zLXut9KXKc+CARIAr+uMSV9q+OuSZPqi
# hqrVuZXn0dtI5Ws3zAztXYDkPm2uExEJ/1QLrY/JTv+2oNxoWe2djXUwAeznJF9C
# 53xZLZZ3XLeZos44vAiAf4BhcKHspBRezV254l7ncYTcz17zwYQWN/Ml088zR8Q2
# TgQ14cqIGMevb3SGgy71wsFkx7MOcASWokzBnSnBbAlFC+JDmNIb+tFWJHHbjhff
# nu1pq7CS1jDOSGUuTLy0FKc75f1w5yXpO2iGiFN5bWaLcv/C6+kgTa+4Wr8esy8c
# RMGfxFH1N/ICrkMTqKOdKNrEXJ0CAwEAAaOCAYYwggGCMB8GA1UdIwQYMBaAFGyK
# lKJ3sYByHYF6Fqry3M5m7kXAMB0GA1UdDgQWBBQX5kvIGkvJp6ZwtExNXsj2NtQw
# mDAOBgNVHQ8BAf8EBAMCAYYwZwYDVR0gBGAwXjALBglghkgBZQIBCyQwCwYJYIZI
# AWUCAQsnMAsGCWCGSAFlAgELKjALBglghkgBZQIBCzswDAYKYIZIAWUDAgEDDTAM
# BgpghkgBZQMCAQMRMAwGCmCGSAFlAwIBAycwEgYDVR0TAQH/BAgwBgEB/wIBADAM
# BgNVHSQEBTADgAEAMDcGA1UdHwQwMC4wLKAqoCiGJmh0dHA6Ly9jcmwuZGlzYS5t
# aWwvY3JsL0RPRFJPT1RDQTMuY3JsMGwGCCsGAQUFBwEBBGAwXjA6BggrBgEFBQcw
# AoYuaHR0cDovL2NybC5kaXNhLm1pbC9pc3N1ZWR0by9ET0RST09UQ0EzX0lULnA3
# YzAgBggrBgEFBQcwAYYUaHR0cDovL29jc3AuZGlzYS5taWwwDQYJKoZIhvcNAQEL
# BQADggEBAAYb1S9VHDiQKcMZbudETt3Q+06f/FTH6wMGEre7nCwUqXXR8bsFLCZB
# GpCe1vB6IkUD10hltI62QMXVx999Qy4ckT7Z/9s4VZC4j1OvsFL5np9Ld6LU+tRG
# uaCblPERLqXOdeq0vgzcgiS+VgxpozEEssYTHLa3rZotnG/cQhr7aA+pVIKh3Q0D
# ZDyhuhGCSj8DTWBt8whxDUUSoGXfNsaFQgfYdzYWdzNbkvmFzrXDrZMHwSihzEPF
# teDSVLwy98Y8i/uStWIuibX+Rt6QL8WUIH/730dw+s8bTuEMv6vKmFtnssiZ0Wvb
# 5tZH41HdkdDZk+jWlIw6YtxGdK4hexUxggITMIICDwIBATBhMFoxCzAJBgNVBAYT
# AlVTMRgwFgYDVQQKEw9VLlMuIEdvdmVybm1lbnQxDDAKBgNVBAsTA0RvRDEMMAoG
# A1UECxMDUEtJMRUwEwYDVQQDEwxET0QgSUQgQ0EtNjMCAxI0bzANBglghkgBZQME
# AgEFAKCBhDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEM
# BgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqG
# SIb3DQEJBDEiBCA/7W38gG0KhBHeWNVpoaJUGywksAv8TucdSfJLUqls3DANBgkq
# hkiG9w0BAQEFAASCAQBo0m+0caIgpl/CLsAVkx/kcTcJhRqGsERj6U7GXnGrgPe9
# DugqRdq1hzHhdfBYRLnOn1Vm5M3QEbXrG+Pk6Cd1mxOAcNfYq9nTXt5sO3Yr9UQ+
# H2YIC/aUDas8M9E/9n0cMUr6/bBbf+zlLFD0q8kbWR9xZP+h5j3fL6wybBFJKNeQ
# 1FAjg9BWBwWmKLdbZNEZIw0ntOEATB2xpCkfTKd4wdIt5DuOdJLiX9y348utbXwZ
# pIOG0F6Wy7MHnWyGUDuY10LbmiV4fDxMDnLWl8J+lBsz562HwNkv9BQaYam0IZXk
# Ju1sU2BX2KJmV2cA/nvICwWYfwJGvy+d/cjr25+D
# SIG # End signature block
