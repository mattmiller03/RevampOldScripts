#PURPOSE: This script will collect all VM inventory from all vCenter servers

#CHANGELOG
#Version 1.00 - 07/11/24 - MDR - Initial version
#Version 1.01 - 08/06/24 - MDR - Modified section where autofilter is set at the end of the script
#Version 1.02 - 08/08/24 - MDR - Add parameter for passing the vCenter password
#Version 1.03 - 09/13/24 - MDR - Pass creds for svc_vrapsh and configure the script to work differently depending on whether run from Orchestrator or manually
#Version 1.04 - 09/16/24 - MDR - Renaming the output file, changing report path directory, moving older versions of this report to SavedFiles folder, renaming $vCenterCreds variable to $vCenterPassword
#Version 1.05 - 10/03/24 - MDR - Importing Base64 passwords from .csv rather than passing as parameters to improve security
#Version 1.06 - 10/04/23 - MDR - Minor Corrections
#Version 1.07 - 10/25/24 - MDR - Added "-FreezePane 2, 2" to all Export-Excel lines
#Version 1.08 - 01/17/25 - MDR - Adding capture of vTPM
#Version 1.09 - 02/03/25 - MDR - Collect Resource Pool
#Version 1.10 - 02/28/25 - MDR - Collect Folder name
#Version 1.11 - 04/21/25 - MDR - Updated with the new Dev-Test vCenter server
#Version 1.12 - 05/13/25 - MDR - Added last backup date and clear variables so templates don't pull data from other scans
#Version 1.13 - 08/11/25 - MDR - Renamed $GuestOS to $ConfiguredGuestOS and then added a new variable $RunningGuestOS.  These collect how the OS is configured in VM settings and the OS VMware Tools sees
#Version 1.14 - 10/09/25 - MDR - Store report to local temp folder before moving it to the network share
#Version 1.15 - 11/21/25 - MDR - Fixing issue where systems with multiple tags are outputting as System.Object[]

Param ( $ParameterPath, $RunFromOrchestrator )

#Version 1.06 - Only import parameters from file if run from Orchestrator
If ($RunFromOrchestrator -eq "True") {
    #Version 1.05 - Import Base64 passwords from CSV
    $Base64Passwords = Import-CSV $ParameterPath

    #Version 1.05 - Delete the CSV file since it is no longer needed and for security reasons
    Remove-Item $ParameterPath | Out-Null

    #Version 1.05 - Store passwords to temp variables
    $vCenterPasswordBase64 = $Base64Passwords.vCenterPassword
    $VRAPasswordBase64 = $Base64Passwords.VRAPassword

    #Version 1.05 - Decode passwords from Base64 to plain text
    $vCenterPassword = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($vCenterPasswordBase64))
    $VRAPassword = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($VRAPasswordBase64))
}

#Configure variables
$TodaysDate = Get-Date -Format "MMddyyyy" #Version 1.04 - Made this a 4 digit year
$ReportPath = "\\orgaze.dir.ad.dla.mil\DCC\VirtualTeam\Reports" #Version 1.04 - Changed path to VPM share \ Reports folder
$ReportFileName = "Master_VMInventory_$TodaysDate.xlsx" #Version 1.04 - Storing report file name in a variable instead of hardcoding throughout the script
$ShortReportName = "Master_VMInventory" #Version 1.04 - Adding a short name which is used when finding older versions of this report to move to SavedFiles
$VMInventoryData = New-Object System.Collections.Generic.List[System.Object]

$ReportColumnHeaders = "Name", "PowerState", "Cluster", "Configured_Guest_OS", "Running_Guest_OS", "Notes", "Floppydrive", "USB Controller", "No USB Cont", "Needs Disk Consolidated", "NumCpu", "CPU Hot Add", "Memory GB", "Memory Hot Add",
                        "Hardware Version", "VMTools Status", "Datacenter", "CD Drive", "CD Connected", "Template", "IP #1", "IP #2", "IP #3", "1st vNic", "2nd vNic", "3rd vNic", "4th vNic", "5th vNic",
                        "VMTools Version", "Disk1", "Disk2", "Disk3", "Disk4", "Disk5", "Disk6", "Disk7", "Disk8", "Disk9", "Disk10", "Disk11", "Disk12", "vCenter", "Firmware", "NestedHVEnabled",
                        "EfiSecureBootEnabled", "VvtdEnabled", "VbsEnabled", "Hostname not equal VMname", "Application_Tag1", "Application_Tag2", "Function_Tag", "OperatingSystem_Tag", "EnclaveID_Tag",
                        "ResourceType_Tag", "Site_Location_Tag", "VlanID_Tag1", "VlanID_Tag2", "VlanID_Tag3", "VlanID_Tag4", "TeamPoc_Tag1", "TeamPoc_Tag2", "TeamPoc_Tag3", "TeamPoc_Tag4", "Change_Number", "vTPM",
                        "ResourcePool", "FolderName", "LastBackup"

#Version 1.14 - Ensure Temp folder exists
If (!(Test-Path "C:\Temp")) {
    New-Item "C:\Temp" -ItemType Directory | Out-Null
}

#Version 1.03 - Create an ouput file for commands being executed
Start-Transcript C:\Temp\Get_All_VM_Inventory_$TodaysDate.txt

Clear-Host

#If a ReportPath drive mapping already exists then remove it
Get-PSDrive -Name ReportPath -ErrorAction SilentlyContinue | Remove-PSDrive -ErrorAction SilentlyContinue

#Version 1.03 - If this is run from Orchestrator then generate a credential to perform the PSDrive mapping
If ($RunFromOrchestrator -eq "True") {
    #Version 1.03 - Create a credential for svc_vrapsh
    $VRACred = New-Object System.Management.Automation.PSCredential -ArgumentList @("DIR\svc_vrapsh",(ConvertTo-SecureString -String $VRAPassword -AsPlainText -Force))
    #Version 1.03 - Map a drive to the ReportPath
    New-PSDrive -Name "ReportPath" -PSProvider FileSystem -Root $ReportPath -Credential $VRACred | Out-Null
} Else { #If not run from Orchestrator then just map the PSDrive with current credentials
    #Version 1.06 - Added a silentlycontinue to prevent reporting when the drive is already mapped
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

#Version 1.02 - If the vCenter password is passed then no need to prompt for it
#Version 1.04 - Changed variable from $vCenterCreds to $vCenterPassword
If ($vCenterPassword -eq $null) {
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

    #The switch will set a few variables based on the vCenter being worked
    #Version 1.15 - Updated all the vCenter names
    Switch ($vCenterServer.ServerName) {
        daisv0pp241.dir.ad.dla.mil {
            $apptagname = "vCenter-Prod-App-name"
            $enclaveidtagname = "vCenter-Prod-Enclave_ID"
            $ostagname = "vCenter-PROD-Os-version"
            $resourcetagname = "vCenter-Prod-Resource_Type"
            $funtagname = "vCenter-Prod-Function"
            $vlanidtagname = "vCenter-Prod-VLAN_ID"
            $locationtagname = "vCenter-Prod-Site_Location"
            $poctagname = "vCenter-Prod-TeamPOC"
        } #End of Dayton Prod vCenter
        trisv0pp241.dir.ad.dla.mil {
            $apptagname = "vCenter-Prod-App-name"
            $enclaveidtagname = "vCenter-Prod-Enclave_ID"
            $ostagname = "vCenter-PROD-Os-version"
            $resourcetagname = "vCenter-Prod-Resource_Type"
            $funtagname = "vCenter-Prod-Function"
            $vlanidtagname = "vCenter-Prod-VLAN_ID"
            $locationtagname = "vCenter-Prod-Site_Location"
            $poctagname = "vCenter-Prod-TeamPOC"
        } #End of Tracy Prod vCenter
        daisv0tp231.dir.ad.dla.mil { #Dayton DEV new tag names working, others need setup and tested
        #  write-host "Working on Dayton DEV Vcenter... " $vcenter
            $apptagname = "vCenter-Dev-App-name"
            $enclaveidtagname = "vCenter-Dev-Enclave_ID"
            $ostagname = "vCenter-Dev-Os-version"
            $resourcetagname = "vCenter-Dev-Resource_Type"
            $funtagname = "vCenter-Dev-Function"
            $vlanidtagname = "vCenter-Dev-VLAN_ID"
            $poctagname = "vCenter-Dev-TeamPOC"
            $locationtagname = "vCenter-Dev-Site_Location"
        } #End of Dayton Dev vCenter
        daisv0pp271.ics.dla.mil {
          #  write-host "Working on Dayton OT Vcenter... " $vcenter
            $apptagname = "vCenter-OT-App-name"
            $enclaveidtagname = "vCenter-OT-Enclave_ID"
            $ostagname = "vCenter-OT-Os-version"
            $resourcetagname = "vCenter-OT-Resource_Type"
            $funtagname = "vCenter-OT-Function"
            $vlanidtagname = "vCenter-OT-VLAN_ID"
            $locationtagname = "vCenter-OT-Site_Location"
            $poctagname = "vCenter-OT-TeamPOC"
        } #End of Dayton OT vCenter
        trisv0pp271.ics.dla.mil {
          #  write-host "Working on Tracy OT Vcenter... " $vcenter
            $apptagname = "vCenter-OT-App-name"
            $enclaveidtagname = "vCenter-OT-Enclave_ID"
            $ostagname = "vCenter-OT-Os-version"
            $resourcetagname = "vCenter-OT-Resource_Type"
            $funtagname = "vCenter-OT-Function"
            $vlanidtagname = "vCenter-OT-VLAN_ID"
            $locationtagname = "vCenter-OT-Site_Location"
            $poctagname = "vCenter-OT-TeamPOC"
        } #End of Tracy OT vCenter
        daisv0pp261.dir.ad.dla.mil {
            $apptagname = "vCenter-VDI-App-name"
            $enclaveidtagname = "vCenter-VDI-Enclave_ID"
            $ostagname = "vCenter-VDI-Os-version"
            $resourcetagname = "vCenter-VDI-Resource_Type"
            $funtagname = "vCenter-VDI-Function"
            $vlanidtagname = "vCenter-VDI-VLAN_ID"
            $locationtagname = "vCenter-VDI-Site_Location"
            $poctagname = "vCenter-VDI-TeamPOC"
        } #End of Dayton VDI 1 vCenter
        daisv0pp262.dir.ad.dla.mil {
            $apptagname = "vCenter-VDI-App-name"
            $enclaveidtagname = "vCenter-VDI-Enclave_ID"
            $ostagname = "vCenter-VDI-Os-version"
            $resourcetagname = "vCenter-VDI-Resource_Type"
            $funtagname = "vCenter-VDI-Function"
            $vlanidtagname = "vCenter-VDI-VLAN_ID"
            $locationtagname = "vCenter-VDI-Site_Location"
            $poctagname = "vCenter-VDI-TeamPOC"
        } #End of Dayton VDI 2 vCenter
        klisv0pp251.dir.ad.dla.mil {
            $apptagname = "vCenter-Kleber-App-name"
            $enclaveidtagname = "vCenter-Kleber-Enclave_ID"
            $ostagname = "vCenter-Kleber-Os-version"
            $resourcetagname = "vCenter-Kleber-Resource_Type"
            $funtagname = "vCenter-Kleber-Function"
            $vlanidtagname = "vCenter-Kleber-VLAN_ID"
            $locationtagname = "vCenter-Kleber-Site_Location"
            $poctagname = "vCenter-Kleber-TeamPOC"
        } #End of Kleber Prod vCenter
        Default { #Version 1.15 - If the vCenter server name is not found then throw an error
            Write-host "`nThe tag section of the script is missing vCenter server $($vCenterServer.ServerName).  This must be resolved" -ForegroundColor Red
        }
    } #End of switch vCenter

    $VMList = Get-View -ViewType VirtualMachine | Sort Name

    #Remove any VMs that aren't connected to vCenter for some reason
    $VMList = $VMList | Where { $_.OverallStatus -ne "gray" }

    #Populate variables for progress bar
    $TotalVMs = $VMList.Count
    $CurrentVMNum = 1

    #Loop through each VM host
    ForEach ($CurrentVM in $VMList) {
        #Version 1.03 - If this is run from Orchestrator then output the ESX host name to get written to the transcript file
        If ($RunFromOrchestrator -eq "True") {
            Write-Host $CurrentVM.Name
        } Else { #If not from Orchestrator then write a progress bar
            #Update progress bar
            Write-Progress -Activity "Getting VM data" -Status "$CurrentVMNum of $TotalVMs" -PercentComplete ($CurrentVMNum / $TotalVMs * 100)
        }

        #Version 1.12 - Clear variables to ensure that templates don't pull values from the previous scan
        Clear-Variable cluster, notes, floppy, usbcontroller, usbcont2, diskconsolidate, toolsstatus, DataCenterName, cddrive, cdstate, vmip1, vmip2, vmip3, vmnic1, vmnic2, vmnic3, vmnic4, vmnic5,`
                       nestedHVEnabled, EfiSecureBoot, VvtdEnabled, VbsEnabled, hostname, funtag, ostag, enclaveidtag, resourcetag, locationtag, chgnum, vTPM, ResourcePool, FolderName, LastBackupDate -ErrorAction SilentlyContinue
        $disk=""
        $apptags=""
        $vlantags=""
        $poctags=""

        #Check if this VM is a template or a regular VM
        If ($CurrentVM.Config.Template -eq "True" ) {
            $template = $CurrentVM.Config.Template
            $vmpwr = $CurrentVM.summary.Runtime.PowerState
            $toolsversion = $CurrentVM.summary.guest.ToolsVersionStatus
            $ConfiguredGuestOS  = $CurrentVM.Config.GuestFullName
            $RunningGuestOS  = $CurrentVM.Guest.GuestFullName #Version 1.13 - Added this line
            $firmware = $CurrentVM.Config.Firmware
            $memhotadd = $CurrentVM.Config.MemoryHotAddEnabled
            $MemoryMB = $CurrentVM.summary.config.memorysizeMB
            $cpuhotadd = $CurrentVM.Config.CpuHotAddEnabled
            $NumCpu = $CurrentVM.summary.config.NumCPU
            $version = $CurrentVM.config.version
        } Else { #Have a VM to process
            #Set template to false
            $template = "False"
            #Need the .NET object view of the VM
            $VMData = Get-VM -Name $CurrentVM.Name
            #Need the VIobjectByView to get all IPs on the VM Guest, disks and ustom Fields for Chg_Number
            $VMView = $CurrentVM |get-VIobjectByVIView

            $VMName = ($VMData).Name
            $vmpwr = $CurrentVM.summary.Runtime.PowerState
            $ConfiguredGuestOS  = $CurrentVM.Config.GuestFullName
            $RunningGuestOS  = $CurrentVM.Guest.GuestFullName #Version 1.13 - Added this line
            $firmware = $CurrentVM.Config.Firmware
            $diskconsolidate = $CurrentVM.summary.runtime.consolidationNeeded
            $NumCpu = $CurrentVM.summary.config.NumCPU
            $cpuhotadd = $CurrentVM.config.CPUHotAddEnabled
            $MemoryMB = $CurrentVM.summary.config.memorysizeMB
            $memhotadd = $CurrentVM.config.MemoryHotAddEnabled
            $version = $CurrentVM.config.version
            $DataCenterName = (Get-Datacenter -VM $CurrentVM.name).Name
            $notes = $CurrentVM.Config.Annotation
            $cluster = (Get-Cluster -VM $VMName).Name
            $floppy = (Get-FloppyDrive -VM $CurrentVM.Name).Name
            $ResourcePool = ($VMData.ResourcePool).Name #Version 1.09 - Collect Resource Pool
            $FolderName = $VMData.Folder.Name #Version 1.10 - Collect Folder Name

            if ($notes.length -ge 1) { #Need to remove the newline char
                if ($notes.Contains("`n")) {
                    $notes = $notes.Replace("`n"," - ")
                } # End of if $notes.contains
            } # End of if $notes.length

            if ($ConfiguredGuestOS -like "Microsoft Windows Server 201*")  { #Get the USB Controller for 2012 and 2016
                $usbcontroller = ($CurrentVM.Config.Hardware.Device | where {$_ -is [VMware.Vim.VirtualUSBXHCIController]} | %{ New-Object PSObject -Property @{Name = $_.DeviceInfo.Label}}).Name
                if (!$usbcontroller) { 
                    $usbcontroller = ($CurrentVM.Config.Hardware.Device | where {$_ -is [VMware.Vim.VirtualUSBController]} | %{ New-Object PSObject -Property @{Name = $_.DeviceInfo.Label}}).Name
                } # End of if controller = "Not Assigned"
            } Else { # NOT Windows 2012 or 2016
                $usbcontroller = ($CurrentVM.Config.Hardware.Device | where {$_ -is [VMware.Vim.VirtualUSBController]} | %{ New-Object PSObject -Property @{Name = $_.DeviceInfo.Label}}).Name
            } #End of IF/ELSE ConfiguredGuestOS test

            If (!$usbcontroller) { 
                $usbcontroller = "No USB Cont"
                If ($ConfiguredGuestOS -like "Microsoft Windows*") {
                    $usbcont2 = "Windows OS w/No USB Controller, need to add"
                } else { 
                    $usbcont2 = "NOT Windows OS, don't add a USB Cont"
                } # End of IF/ESLE ConfiguredGuestOS test
            } else {
                $usbcont2 = "False"
            }  # End of the usbcontroller section

            $cd = Get-CDDrive -VM $CurrentVM.name
            if ($cd.Count -eq 0) {
                [string]$cddrive = "No CDROM drive"
                [string]$cdstate = "FALSE" 
            } else {
                $cddrive = $cd.name
                $cdstate = $cd.ConnectionState.Connected 
            } # End of if cd = null statement
            [string]$vmip = $VMView.guest.IPAddress
            $vmip1,$vmip2,$vmip3 = $vmip.split(" ",3) # Split the multiple IPs into seperate variables
            #NOTE: May need to put a test for vm's vCLS and skip those
            $vNics = Get-NetworkAdapter -VM $CurrentVM.name | select -Expand Type
            #Need to test vNics for a null value, don't process if null            	
            if ($vNics.count -gt 0) { 
                [string]$vmnics = $vNics
                $vmnic1,$vmnic2,$vmnic3,$vmnic4,$vmnic5 = $vmnics.split(' ',5) # Split the multiple nics into separate variables
            } else {
                $vmnic1 = "No Network device"
            }
            $toolsstatus = $CurrentVM.summary.guest.ToolsRunningStatus
            if ($toolsstatus -like "guest*") {
                $toolsstatus = $toolsstatus.TrimStart("guest")
            }
            #NOTE: Some VMs need to be marked as Guest Managed vs the version number
            $toolsversionstatus = $CurrentVM.summary.guest.ToolsVersionStatus
            Switch ($toolsversionstatus) {
                guestToolsNotInstalled {
                    $toolsversion = "VMTools Not Installed"
                }
                guestToolsUnmanaged {
                    $toolsversion = "Guest Managed"
                }
                guestToolsCurrent {
                    $toolsversion = (Get-VMGuest -VM $CurrentVM.name).ToolsVersion
                }
                guestToolsNeedUpgrade {
                    $toolsversion = (Get-VMGuest -VM $CurrentVM.name).ToolsVersion
                }
                guestToolsNotInstalled {
                    $toolsversion = "VMTools Not Installed"
                }
            } #end of Switch                        
            if ($vmpwr -eq "poweredOff" -And $ConfiguredGuestOS -like "*Windows*" ) {
                $toolsversion = "Windows VM Powered Off"
            }
            $disk = @() #Initialize the array for the next VM
            $i = 0 #Being used to cycle thru disks
            $vmdisk = $VMView | get-harddisk 
            $vmdiskcount = $vmdisk.Count
            foreach ($vm_d in $vmdisk) {
                $disk += [math]::Round($vm_d.CapacityGB)
                $i = $i + 1   
            } #End of foreach VMDisk loop
            $firmware = $VMView.ExtensionData.Config.Firmware
            $nestedHVEnabled = $VMView.ExtensionData.Config.NestedHVEnabled
            $EfiSecureBoot = $VMView.ExtensionData.Config.BootOptions.EfiSecureBootEnabled
            $VvtdEnabled = $VMView.ExtensionData.Config.flags.VvtdEnabled
            $VbsEnabled = $VMView.ExtensionData.Config.flags.VbsEnabled
            $hostname = $VMView.Guest.Hostname
            if ($hostname -ne $null) { # Have a hostname from the vm's guest OS
                $ShortName = $hostname.split(".")[0]
                if ($CurrentVM.name -eq $ShortName) {
                    # Majority of hostnames match vmname, so don't display matches just different names
                    $hostname = ""
                } #End of if vmname -eq hostname
            } else { # Don't have a hostname from the vm, numerous reasons
                $hostname = "NO hostname returned from vm"
            } # End of else No hostname returned from VM

            #TAG Processing
            $apptags = @()
            #NOTE: Problem with Kleber returning multiple tags for each Category, had to change things
            #  Once tags Function, ResourceType, Application, Enclave were remove things started working again,
            #  added these four new tags back in using PowerCLI commands and extensive testing things worked again
            $apptag = (Get-TagAssignment -Entity $VMData -Category $apptagname).Tag.Name
            foreach ($apptag_d in $apptag) {
                if ($apptag_d -ne $null) {
                    $apptags += $apptag_d
                } else {
                    $apptags[0] = ""
                    $apptags[1] = ""
                }
            } #End of foreach apptag loop

            $funtag = (Get-TagAssignment -Entity $VMData -Category $funtagname).Tag.Name
            $ostag = (Get-TagAssignment -Entity $VMData -Category $ostagname).Tag.Name
            $enclaveidtag = (Get-TagAssignment -Entity $VMData -Category $enclaveidtagname).Tag.Name
            $resourcetag = (Get-TagAssignment -Entity $VMData -Category $resourcetagname).Tag.Name
            $locationtag = (Get-TagAssignment -Entity $VMData -Category $locationtagname).Tag.Name
            
            #Noticed a new tag being added 4/7/23, one VM with multiple VLANs
            $vlantags = @() 
            $vlantag = Get-TagAssignment -Entity $VMData -Category $vlanidtagname | select Tag
            $vmnetworktype = $VMView.ExtensionData.Network |select Type

            if ($VMData.name -like "vCLS*" -And $vlantag -eq $null) {
                $vlantags += "No_Network"
            } elseif ($vlantag -eq $null -and $vmnetworktype -like "*Network*") {
                #No VLAN Tag assigned and on a standard switch
                $vlantags += "Std_Switch"
            } elseif ($vmnetworktype -like "*Distributed*") {
                if ($vlantag -ne $null) {
                    foreach ($vlantag_d in $vlantag) {
                        if ($vlantag_d.tag.count -gt 0) {
                            #write-host "VLANTAG_D has data and is NOT NULL"
                            [string]$vlantag1 = $vlantag_d
                                $cutthis = '@{'
                                $vlantag2 = $vlantag1.TrimStart($cutthis)
                                #$vlantag1 = $vlantag2.Replace("Tag=VLAN ID/","")
                                $vlantag1 = $vlantag2.Replace("Tag=","")
                                $vlantag2 = $vlantag1.Replace($vlanidtagname,"")
                                $vlantag1 = $vlantag2.Replace("/","")
                                $vlantag_d = $vlantag1.Replace("}","")
                                $vlantags += $vlantag_d
                        } else {
                            $vlantags = "None"
                        } # End of IF/ELSE vlantag_d -gt 0
                    } #End of foreach vlantag
                } else {  
                    $vlantags += "None"                 
                #write-host -ForegroundColor Red "No VLAN Assigned to... " $vm.name $vlantags
                } #End of IF/ELSE vlantag -gt 0
            }#End of if vm.network.type Distributed
            elseif ($vlantag.tag.count -gt 0 -and $vmnetworktype -like "*Network*")  { #Have an existing vlantag on a Std Sw
            #  write-host -ForeGroundColor Yellow "Have a STD Switch with... " $vlantag
                #Need to strip off the $vlantagidname from the object
                if ($vlantag.tag.count -eq 1) {
                # write-host -ForeGroundColor Yellow "This is the vlantag... " $vlantag
                    [string]$temp = $vlantag
                        $cutthis = "@{Tag="
                        $temp = $temp.TrimStart($cutthis)
                        $cutthis = $vlanidtagname
                        $temp = $temp.TrimStart($cutthis)
                        $temp = $temp.Replace("/","")
                        $vlantag = $temp.Replace("}","")
                        $vlantags = @($vlantag)
                #   write-host -ForeGroundColor Cyan "This is the vlantag... " $vlantag
                } #End of IF vlantag.tag.count = 1
                elseif ($vlantag.tag.count -gt 1) {
                    foreach ($vlantag_d in $vlantag) {
                        if ($vlantag_d.tag.count -gt 0) {
                            #write-host "VLANTAG_D has data and is NOT NULL"
                            [string]$vlantag1 = $vlantag_d
                                $cutthis = '@{'
                                $vlantag2 = $vlantag1.TrimStart($cutthis)
                                #$vlantag1 = $vlantag2.Replace("Tag=VLAN ID/","")
                                $vlantag1 = $vlantag2.Replace("Tag=","")
                                $vlantag2 = $vlantag1.Replace($vlanidtagname,"")
                                $vlantag1 = $vlantag2.Replace("/","")
                                $vlantag_d = $vlantag1.Replace("}","")
                                $vlantags += $vlantag_d
                        } # End of IF vlantag count > 1
                    } #End of foreach vlantag
                } #End of else/if vlantag.tag.count > 1
                elseif ($vlantag.tag.length -eq 0) {
                    $vlantags += "None"
                } #End of if/else vlantag > 0
            # write-host -ForegroundColor Yellow "This is the VLANTAG on... " $VMData.name $vlantags 
            } #End of ElseIF vm.network loop
            else { #No VLAN or StdSwitch
            #   write-host -ForeGroundColor RED "No VLAN to process"
                $vlantags = ""
            } #End of IF/ELSEIF...ELSE for VLAN tags

            $poctags = @() 
            $poctag = Get-TagAssignment -Entity $VMData -Category $poctagname | select Tag
            if ($poctag.tag.count -gt 0){
                foreach ($poctag_d in $poctag) {
                if ($poctag_d -ne $null) {
                        [string]$poctag1 = $poctag_d
                            $cutthis = '@{'
                            $poctag2 = $poctag1.TrimStart($cutthis)
                            $poctag1 = $poctag2.Replace("Tag=","")
                            $poctag2 = $poctag1.Replace($poctagname,"")
                            $poctag1 = $poctag2.Replace("/","")
                            $poctag_d = $poctag1.Replace("}","")
                        $poctags += $poctag_d
                    } # End of IF poctag_d not null
                    else {
                        $poctags = "None"
                        # $poctags[0] = "None Assigned"
                        # $poctags[1] = ""
                        # $poctags[2] = ""
                        # $poctags[3] = "" 
                    } #End of IF/ELSE poctag_d -ne null    
                } #End of foreach poctag loop
            } #End of IF poctag.tag.count -gt 0
            else {                   
                #write-host "No POC Assigned to... " $VMName
                $poctags += "None"
                #write-host "The value of poctag[0] is... " $poctags[0]
            } #End of IF/ELSE poctag.tag.count -gt 0  
            
        #write-host -foregroundcolor yellow "This is the network being processed..." $vlantag $vmnetworktype 

            $array = @() #Need an array as the CustomFields has multiple key/value pairs
            $chgnum = $VMData.CustomFields
            [string]$temp = $chgnum
                $pattern = '(?<=\[).+?(?=\])' #This is to split things between the [ and ]
                $array = ($temp | select-string $pattern -AllMatches).Matches.Value #Put all CustomFields in the Array
            if ($array[0].length -gt 1) {
                $cutthis = "CHG_Number," #Just to get the actual CHG_Number
                $chgnum = $array[0].TrimStart($cutthis)
            } else {
                $chgnum = "None"
            }

            #Version 1.12 - Get the last snapshot date
            $LastBackupInfo = ($VMData.CustomFields | Where { $_.Key -eq "NB_LAST_BACKUP" }).Value
            #If there is a last backup found
            If ($LastBackupInfo) {
                $LastBackupSplitComma = $LastBackupInfo -Split ","
                $LastBackupDateLong = $LastBackupSplitComma[0].Substring(0, $LastBackupSplitComma[0].Length-6)
                $LastBackupDate = $LastBackupDateLong -replace '^\S+\s*','' #Remove the first word from the backup date.  This word is the day of the week like "Mon"
            }

            #Version 1.08 - Capture whether a vTPM is installed
            If (Get-VTPM -VM $CurrentVM.Name) {
                $vTPM = "True"
            } Else {
                $vTPM = "False"
            }
        } #End of if/else Template test, processed a VM
        
        #Store data from the VM host
        $VMInventoryData.add((New-Object "psobject" -Property @{"Name"=$CurrentVM.Name;
        "PowerState"=$vmpwr;"Cluster"=$cluster;"Configured_Guest_OS"=$ConfiguredGuestOS;"Running_Guest_OS"=$RunningGuestOS;"Notes"=$notes;"Floppydrive"=$floppy;"USB Controller"=$usbcontroller;"No USB Cont"=$usbcont2;"Needs Disk Consolidated"=$diskconsolidate;
        "NumCpu"=$NumCpu;"CPU Hot Add"=$cpuhotadd;"Memory GB"=$MemoryMB;"Memory Hot Add"=$memhotadd;"Hardware Version"=$version;"VMTools Status"=$toolsstatus;"Datacenter"=$DataCenterName;"CD Drive"=$cddrive;
        "CD Connected"=$cdstate;"Template"=$template;"IP #1"=$vmip1;"IP #2"=$vmip2;"IP #3"=$vmip3;"1st vNic"=$vmnic1;"2nd vNic"=$vmnic2;"3rd vNic"=$vmnic3;"4th vNic"=$vmnic4;"5th vNic"=$vmnic5;
        "VMTools Version"=$toolsversion;"Disk1"=$disk[0];"Disk2"=$disk[1];"Disk3"=$disk[2];"Disk4"=$disk[3];"Disk5"=$disk[4];"Disk6"=$disk[5];"Disk7"=$disk[6];"Disk8"=$disk[7];"Disk9"=$disk[8];
        "Disk10"=$disk[9];"Disk11"=$disk[10];"Disk12"=$disk[11];"vCenter"=$vCenterServer.ServerName;"Firmware"=$firmware;"NestedHVEnabled"=$nestedHVEnabled;"EfiSecureBootEnabled"=$EfiSecureBoot;"VvtdEnabled"=$VvtdEnabled;
        "VbsEnabled"=$VbsEnabled;"Hostname not equal VMname"=$hostname;"Application_Tag1"=$apptags[0];"Application_Tag2"=$apptags[1];"Function_Tag"=$funtag;"OperatingSystem_Tag"=$ostag;
        "EnclaveID_Tag"=$enclaveidtag;"ResourceType_Tag"=$resourcetag;"Site_Location_Tag"=$locationtag;"VlanID_Tag1"=$vlantags[0];"VlanID_Tag2"=$vlantags[1];"VlanID_Tag3"=$vlantags[2];"VlanID_Tag4"=$vlantags[3];
        "TeamPoc_Tag1"=$poctags[0];"TeamPoc_Tag2"=$poctags[1];"TeamPoc_Tag3"=$poctags[2];"TeamPoc_Tag4"=$poctags[3];"Change_Number"=$chgnum;"vTPM"=$vTPM;"ResourcePool"=$ResourcePool;"FolderName"=$FolderName;
        "LastBackup"=$LastBackupDate}))
    
        #Increment progress bar counter
        $CurrentVMNum++
    }
    
    #Close the progress bar
    Write-Progress -Activity "Getting VM data" -Completed
}

Write-Host "`nWriting data to Excel" -ForegroundColor Cyan

#Version 1.04 - Move the old reports into the SavedFiles path
Get-ChildItem -Path $ReportPath -Filter "$ShortReportName*" | ForEach {
    Move-Item -Path $_.FullName "$ReportPath\SavedFiles" -Force
}

#Set the data headers in the correct order
$FinalOutput = $VMInventoryData | Select $ReportColumnHeaders | Sort Name

#All_VMs
$FinalOutput | Export-Excel -WorksheetName "All_VMs" -Path "C:\Temp\$ReportFileName" -AutoSize -Append -FreezePane 2, 2
#Missing Tags
$FinalOutput | Where { $_.Name -notlike "vCLS-*" -and ( $_.OperatingSystem_Tag -eq "None" -or $_.Site_Location_Tag -eq "None" -or $_.VlanID_Tag1 -eq "None" -or $_.TeamPoc_Tag1 -eq "None") }`
                | Export-Excel -WorksheetName "MissingTags" -Path "C:\Temp\$ReportFileName" -AutoSize -Append -FreezePane 2, 2
#VMs_BIOS
$FinalOutput | Where { $_.Firmware -eq "bios" } | Export-Excel -WorksheetName "VMs_BIOS" -Path "C:\Temp\$ReportFileName" -AutoSize -Append -FreezePane 2, 2
#VMs_PoweredOff
$FinalOutput | Where { $_.PowerState -eq "poweredOff" } | Export-Excel -WorksheetName "VMs_PoweredOff" -Path "C:\Temp\$ReportFileName" -AutoSize -Append -FreezePane 2, 2

#Create a hash array list for vCenter Servers
$vCenterServerList = @("daisv0pp261.dir.ad.dla.mil","daisv0pp262.dir.ad.dla.mil","daisv0tp231.dir.ad.dla.mil","daisv0pp241.dir.ad.dla.mil","daisv0pp271.ics.dla.mil","klisv0pp251.dir.ad.dla.mil","trisv0pp241.dir.ad.dla.mil","trisv0pp271.ics.dla.mil")
$ExcelTabNameList = @("daisv0pp261","daisv0pp262","DaytonDev","DaytonProd","DaytonICS","Kleber","TracyProd","TracyICS")

#Loop through each vCenter Server and create a tab for them
For ($CurrvCenter = 0; $CurrvCenter -lt $vCenterServerList.Count; $CurrvCenter++) {
    $FinalOutput | Where { $_.vCenter -eq $vCenterServerList[$CurrvCenter] } | Export-Excel -WorksheetName $ExcelTabNameList[$CurrvCenter] -Path "C:\Temp\$ReportFileName" -AutoSize -Append -FreezePane 2, 2
}

#CPU_HotAdd_FALSE
$FinalOutput | Where { $_."CPU Hot Add" -eq $false } | Export-Excel -WorksheetName "CPU_HotAdd_FALSE" -Path "C:\Temp\$ReportFileName" -AutoSize -Append -FreezePane 2, 2
#Memory_HotAdd_FALSE
$FinalOutput | Where { $_."Memory Hot Add" -eq $false } | Export-Excel -WorksheetName "Memory_HotAdd_FALSE" -Path "C:\Temp\$ReportFileName" -AutoSize -Append -FreezePane 2, 2
#VMToolsVersion
#Get a list of all numeric version numbers
$VMToolsVersionList = $FinalOutput."VMTools Version" | Select -Unique | Where { $_ -match '^[0-9]+$*' } | ForEach { [Version]$_ }
[String]$VMwareToolsLatestVersion = $VMToolsVersionList | Sort -Descending | Select -First 1
$FinalOutput | Where { $_."VMTools Version" -ne $VMwareToolsLatestVersion -and $_."VMTools Version" -match '^[0-9]+$*' } | Export-Excel -WorksheetName "VMToolsVersion" -Path "C:\Temp\$ReportFileName" -AutoSize -Append -FreezePane 2, 2
#FloppyDrives
$FinalOutput | Where { $_.Floppydrive -ne $null } | Export-Excel -WorksheetName "FloppyDrives" -Path "C:\Temp\$ReportFileName" -AutoSize -Append -FreezePane 2, 2

#Version 1.14 - Only do auto filter if not running from the PSH host
If ($RunFromOrchestrator -ne "True") {
    #Set auto filter on all worksheets
    #Open the spreadsheet that was created
    $ExcelPkg = Open-ExcelPackage -Path "C:\Temp\$ReportFileName"
    #Loop through each worksheet
    ForEach ($WorkSheet in $ExcelPkg.Workbook.Worksheets) {
        #Get the range of data in the worksheet
        $UsedRange = $WorkSheet.Dimension.Address

        #Version 1.01 - Added a check just in case the worksheet is blank and doesn't need a filter
        If ($UsedRange -ne $null) {
            #Enable auto filter on that range
            $WorkSheet.Cells[$UsedRange].AutoFilter = $true
        }
    }
    #Close the spreadsheet
    Close-ExcelPackage $excelPkg
}

#Version 1.14 - Move the report to the report folder
Move-Item -Path "C:\Temp\$ReportFileName" -Destination "ReportPath:\"

Write-Host "`nScript Complete.  Report written to $ReportPath\VMinventory-$TodaysDate.xlsx" -ForegroundColor Green

#Version 1.03 - Stop the transcript
Stop-Transcript
# SIG # Begin signature block
# MIIL6gYJKoZIhvcNAQcCoIIL2zCCC9cCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBKrJ4Lvrw7mfpF
# mGCp5A9GEGOYYa2Nvro0nszmD8IdHqCCCS0wggRsMIIDVKADAgECAgMSNG8wDQYJ
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
# SIb3DQEJBDEiBCBqcMP6qcQSu18uNrGdvPeSVmeE55mb03odLanQVBXItzANBgkq
# hkiG9w0BAQEFAASCAQBc/oYAuXx1k773tCcaxwy7t5RxSqiTYY0BhirzufHlqR5s
# U5f/nVU84VaWO/MOMX4RGvcaDxtsqK09FcZv/FnrJqhTUW0Hz6FmcMwRv+mNX2Qo
# twTZOvUZUmsO3fYGK2Q9aynTPjvWml5BVAg0ZrAsU87LVosBX/JfzBWLR2NBg4Mi
# +Bq3HxycOsiLFIpLRC/C36fh9EF2wQ4j+2/VTsuLJYv7DTHbFMc/yPuJ4Ks+YwIQ
# ufLkpeb5NWd1qVWmo5uAPNBrAIRcKXSlO6qYK7uJBbEKsz94bCJbNpC3XFbTpwBA
# vApxCGCcVamIEyrD6omKi+W+eIQOGIl2BDKozPuc
# SIG # End signature block
