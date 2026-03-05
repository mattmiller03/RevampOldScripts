#PURPOSE: This script will collect the Azure VM Inventory Report

<#CHANGELOG
    Version 1.00 - 09/24/24 - JN - Initial version
    Version 1.01 - 10/04/24 - MDR - Minor revisions
    Version 1.02 - 10/08/24 - JN - Minor revisions
    Version 1.03 - 10/23/24 - JN - Removing the Windows_Agent_Old sheet
    Version 1.04 - 02/18/25 - JN - Update to newer Tag Values for App-name and Nic tag values 
    Version 1.05 - 10/27/25 - MDR - Corrected section where if there are multiple IPs on the NIC it will now capture if there is only 1 where it used to only capture if greater than 1
    Version 1.06 - 11/04/25 - MDR - Rewrote the whole NIC data collection section to clearify what is happening there and resolve an issue where IPv6 addresses were not being collected
    Version 1.07 - 11/05/25 - MDR - Because of the change to how NIC data is collected, I needed to move where the "sort" command is so NIC 1 is the first followed by NIC 2 and so on
#>
Param ( $ParameterPath, $ApplicationID, $TenantID, $RunFromOrchestrator )

#Version 1.01 - Only import parameters from file if run from Orchestrator
If ($RunFromOrchestrator -eq "True") {
    #Import Base64 passwords from CSV to improve security
    $Base64Passwords = Import-CSV $ParameterPath

    #Delete the CSV file since it is no longer needed and for security reasons
    Remove-Item $ParameterPath | Out-Null

    #Store passwords to temp variables
    $VRAPasswordBase64 = $Base64Passwords.VRAPassword
    $AzurePasswordBase64 = $Base64Passwords.AzurePassword #Version 1.01 - Corrected to AzurePassword

    #Decode passwords from Base64 to plain text
    $VRAPassword = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($VRAPasswordBase64))
    $AzurePassword = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($AzurePasswordBase64))
}

#Configure variables
$TodaysDate = Get-Date -Format "MMddyyyy"
$ReportPath = "\\orgaze.dir.ad.dla.mil\DCC\VirtualTeam\Reports"
$ReportFileName = "Master_Azure_VMInventory_$TodaysDate.xlsx"
$ShortReportName = "Master_Azure_VMInventory"
$AzureVMInventoryList = New-Object System.Collections.Generic.List[System.Object]

#Report column headers
#Version 1.02 - Added three additional fields, SecurityType, SecureBoot and vTPMStatus
#Version 1.04 - Updated App-name tag value, adding NIC tag value
$ReportColumnHeaders = "VMName","ResourceGroupName","VMStatus","Location","App-name","Function","Impact","Mission",
        "LicenseType","VMSize","ManagedDisk","VM_Memory","VM_Cores","VMAgentVersion","VM_Agent_Status","OSType",
        "OSDisk","OSDiskSize","BootDiag","BootDiagSA","Shutdown","Startup","Role","Role_Assignment","VM_Generation",
        "Nic_1_Name","Nic_1_Tag","Nic_1_AcceleratedNetworking","Nic_1_Subnet","Nic_1_IP1","Nic_1_IP2",
        "Nic_2_Name","Nic_2_Tag","Nic_2_AcceleratedNetworking","Nic_2_Subnet","Nic_2_IP1","Nic_2_IP2",
        "Nic_3_Name","Nic_3_Tag","Nic_3_AcceleratedNetworking","Nic_3_Subnet","Nic_3_IP1","Nic_3_IP2","OS_name","OS_Version","SubscriptionName",
        "SecurityType","SecureBoot","vTPMStatus"

Start-Transcript C:\Temp\Get_Azure_VM_Inventory_$TodaysDate.txt

Clear-Host

#If a ReportPath drive mapping already exists then remove it
Get-PSDrive -Name ReportPath -ErrorAction SilentlyContinue | Remove-PSDrive

#If this is run from Orchestrator then generate a credential to perform the PSDrive mapping
If ($RunFromOrchestrator -eq "True") {
    #Convert the VRA password to a secure string
    $VRACred = New-Object System.Management.Automation.PSCredential -ArgumentList @("DIR\svc_vrapsh",(ConvertTo-SecureString -String $VRAPassword -AsPlainText -Force))
    #Map a drive to the ReportPath
    New-PSDrive -Name "ReportPath" -PSProvider FileSystem -Root $ReportPath -Credential $VRACred | Out-Null
} Else { #If not run from Orchestrator then just map the PSDrive with current credentials
    New-PSDrive -Name "ReportPath" -PSProvider FileSystem -Root $ReportPath -ErrorAction SilentlyContinue | Out-Null
}

If (!(Get-PSDrive -Name "ReportPath")) {
    Write-Host "Failed to connect to ReportPath.  Exiting" -ForegroundColor Red
    Break
}

#If for some reason an Access Denied error happens here then exit the script
Try {
    #If this report exists already then delete it
    If (Test-Path "ReportPath:\$ReportFileName" -ErrorAction Stop) {
        Remove-Item "ReportPath:\$ReportFileName"
    }
} Catch { #If an Access Denied error occurs or any other error trying to reach the $ReportPath
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

#If ImportExcel is not found then prompt for folder Where it is located
If (!$CheckForImportExcel) {
    Write-Host "The ImportExcel module is required for this script to run" -ForegroundColor Red
    Write-Host "`nA copy of this module is located in \\orgaze.dir.ad.dla.mil\J6_INFO_OPS\J64\J64C\WinAdmin\VulnMgt\Software\ImportExcel"
    Write-Host "`nPlace a copy of this module in C:\Program Files\WindowsPowerShell\Modules"
    Break
}

#If this is run from Orchestrator then you'll need to log into the server to configure this
If ($RunFromOrchestrator -eq "True") {
    #Convert the Azure password to a secure string
    $AzureAppCred = New-Object System.Management.Automation.PSCredential -ArgumentList @($ApplicationID, (ConvertTo-SecureString -String $AzurePassword -AsPlainText -Force))
    #Connect to Azure using the supplied creds from Orchestrator
    Connect-AzAccount -Environment AzureUSGovernment -ServicePrincipal -TenantId $TenantID -Credential $AzureAppCred | Out-Null

    #Get a list of subscriptions
    #Version 1.01 - Perform a Try / Catch to avoid reporting errors if this fails
    Try {
        #When testing do just one Subscription
        #$Subscriptions = Get-AzSubscription | where {$_.Name -like "*DADE*" } | Sort Name
        $Subscriptions = Get-AzSubscription | Sort Name
    } Catch { }

    #Check to see if the Orchestrator creds allowed for subscriptions to be read
    If ($Subscriptions.count -lt 1)  {
        Write-Host "`nNot connected to Azure.  Exiting script"
        Break
    } Else { #Notify that Azure is now connected
        Write-Host "`nConnected to Azure US Government" -ForeGroundColor Cyan 
    }
} Else { #If this is run manually then allow for picking the Azure account to login with
    #Get a list of subscriptions, just doing one Subscription to test
    #Version 1.01 - Perform a Try / Catch to avoid reporting errors if this fails
    Try {
        #$Subscriptions = Get-AzSubscription | where {$_.Name -like "*HUB*" } | Sort Name
        $Subscriptions = Get-AzSubscription | Sort Name
    } Catch { }

    #Check to see if subscriptions were collected
    If ($Subscriptions.count -lt 1)  {
        Write-Host "`nNot connected to Azure.  When the popup appears please login to Azure" -ForeGroundColor Yellow
        Connect-AzAccount -EnvironmentName AzureUSGovernment | Out-Null
        #Get a list of subscriptions
        #Version 1.01 - Perform a Try / Catch to avoid reporting errors if this fails
        Try {
            #$Subscriptions = Get-AzSubscription | where {$_.Name -like "*DADE*" } | Sort Name
            $Subscriptions = Get-AzSubscription | Sort Name
            write-host "The Subscriptions being processed are... " $Subscriptions
        } 
        Catch { }
        #Check to see if the supplied creds allowed for subscriptions to be read
        If ($Subscriptions.count -lt 1)  {
            Write-Host "`nNot connected to Azure.  Exiting script"
            Break
        } Else { #Notify that Azure is now connected
            Write-Host "`nConnected to Azure US Government" -ForeGroundColor Cyan 
        }
    }
}

#Loop through each subscription
ForEach ($CurrSubscription in $Subscriptions) {
    #Get subscription name
    $SubscriptionName = $CurrSubscription.Name
    #Get the subscription short name
    $SubscriptionShortName = $SubscriptionName.split('-')[1]
    #Display subscription name
    Write-Host "`nProcessing Subscription: $SubscriptionName"
    #Select the subscription
    Set-AzContext -Subscription $SubscriptionName | Out-Null

    $VMS = Get-AzVM | sort Name
    #Loop through all VMs #End of foreach VM in VMS
#} # End of ForEach Subscriptions
    ForEach ($VM in $VMS) {  
        $date = Get-Date -Format "HH:mm"
        $vm_name = $vm.Name
        write-host "Processing... " $vm_name $date
        $vm_rg = $vm.ResourceGroupName
        If ($vm.StorageProfile.OSDisk.ManagedDisk -like '') {$managed = "No"} Else {$managed = "Yes"}
        $vmsize = $vm.HardwareProfile.VMsize
        #the vmdata var contains a few different values, NumberOfCores, MemoryInMB, OSDiskSizeInMB            
        $vmdata = Get-AzVMSize -Location $vm.Location | where {$_.Name -eq $vmsize}
        $OSType = $vm.StorageProfile.OsDisk.OsType
        $OSDisk = $vm.StorageProfile.OsDisk.Name
        $OSDiskSize = $vm.StorageProfile.OsDisk.DiskSizeGb
        
        #Not all VMs have each tag being pulled, need to test and if null make the tag NONE
        #Version 1.04 - Updated App-name tag value
        $application = $vm.Tags."App-name"
        if ($application -eq $null) {
            $application = "None"
        }

        $function = $vm.tags.Function
        if ($function -eq $null) {
            $function = "None"
        }
        
        $impact = $vm.Tags."Impact-level"
        if ($impact -eq $null) {
            $impact = "None"
        }
        
        $mission = $vm.Tags.Mission
        if ($mission -eq $null) {
            $mission = "None"
        }

        $statuses = (Get-AzVM -ResourceGroupName $vm_rg -Name $($vm_name) -Status)
        $OS_Name = $statuses.OsName
        $OS_Version = $statuses.OsVersion
        #This loop is to get the current VM State
        foreach ($status in $statuses.Statuses) {
            if ($status.Code -notlike "*provisioning*") {
                $vmstatus = $status.Code.Split('/')[1]
                if ($status -like "deallocated") {
                    $vmstaus = "stopped"
                } #End of if -like deallocated
            } #End of if -notlike provisioning
        } #End of foreach status
        $generation = $statuses.HyperVGeneration
        $agent = $statuses.VMAgent.VMAgentVersion
        if ($agent -like "") {
            $agentstatus = "Unknown"
        } #End of if agent -like ""
        else {
            $agentstatus = $statuses.VMAgent.Statuses[0].DisplayStatus
        }#End of if/else agent 
        $bd = $statuses | Select-Object  @{Name="BootDiag"; Expression={$_.BootDiagnostics.ConsoleScreenshotBlobUri}}
        if ($bd -like "@{BootDiag=}") {
            #$bootdiag = $bd
            $bd = "none"
            $sa = "not configured"
        }
        else {  #Need an array to get the boot diag
            $temparray = $bd -split "/"
            $sa = $temparray[2]
            $bd = $temparray[4].Replace("}","")
        } #End of if/else  bootdiag
        
        #Try getting the Role Assignment for the VM
        $vmrole = ""
        $Namearray = @()
        $RoleArray = @()
        $array1 = @()
        $array1 = Get-AzRoleAssignment | where {$_.Scope -eq $vm.Id}

        if ($array1.DisplayName.Count -gt 1) { #This is for VMs with multiple roles assigned
            $count = $array1.Count
            while ($count -gt 0 )  { 
                $temp = $array1.DisplayName[$count - 1] ; $cutthis = '@{' ; $temp = $temp.TrimStart($cutthis) ; $temp = $temp.Replace("DisplayName=","")
                $Namearray += $temp.Replace("}","")
                $temp = $array1.RoleDefinitionName[$count - 1] ; $cutthis = '@{' ; $temp = $temp.TrimStart($cutthis) ; $temp = $temp.Replace("RoleDefinitionName=","")
                $Rolearray += $temp.Replace("}","")
                if ($count -gt 1) { #Multiple roles, more to process
                    $Namearray += ";"
                    $RoleArray += ";"
                } #Process each role individually until one is left to process
                else { #Multiple roles this should be the last to process, assign array to the outpur variable
                    [string]$vmDisplayName += $Namearray
                    [string]$vmRole += $RoleArray
                } #End of the if/else to process all the roles on a VM with multiple roles assigned
                $count = $count - 1
                $vmDisplayName = $Namearray
                $vmRole = $RoleArray
                } # End of do/while loop, clear the arrays just to be safe
                $Namearray = @()
                $Rolearray = @()
        } # End of if array.count > 1
        elseif ($array1.DisplayName.Count -eq 1) { #This is for VMs with only 1 role assigned
                $temp = $array1.DisplayName ; $cutthis = '@{' ; $temp = $temp.TrimStart($cutthis) ; $temp = $temp.Replace("DisplayName=","")
                $vmDisplayName = $temp.Replace("}","")
                $temp = $array1.RoleDefinitionName ; $cutthis = '@{' ; $temp = $temp.TrimStart($cutthis) ; $temp = $temp.Replace("RoleDefinitionName=","")
                $vmRole = $temp.Replace("}","")
        } # End of if array1.count = 1
        elseif ($array1.DisplayName.Count -eq 0 )    { #No role on the VM
                #  write-host "This VM has NO roles...- " $array1 " -..." $array1.DisplayName.count
                $vmDisplayName = "none"
                $vmRole = "none"
            } #End of if for a VM with NO roles assigned on the VM object
            $array1 = @()

        ###Version 1.06 - Rewrote this section to clean it up###

        #Get all NIC IDs
        $NICIDList = $vm.NetworkProfile.NetworkInterfaces.Id
        #Get info for all the NICs with the IDs that were just collected
        #Version 1.07 - Sort NICs so "NIC" will be first followed by "NIC2" then "NIC3" and so on
        #Version 1.08 - Instead of sorting by name, sorting by Primary first
        $NICInfo = Get-AzNetworkInterface | Where { $NICIDList -Contains $_.Id } | Sort Primary -Descending

        ###Store data for NIC 1###
        $nic_1_Name = $NICInfo[0].Name
        $nic_1_Tag = $NICInfo[0].Tag."App-name"
        #Try puting the Accelerated Network setting below this note
        $nic_1_AcceleratedNetworking = $NICInfo[0].EnableAcceleratedNetworking
        #Store the primary IP as IP1
        $nic_1_id_ip1 = ($NICInfo[0].IpConfigurations | Where { $_.Primary -eq $True }).PrivateIpAddress
        $nic_1_id_subnet = $NICInfo[0].IpConfigurations.Subnet.ID | Split-Path -Leaf | Select -First 1
        #If there are multiple IPs on a NIC
        If ($NICInfo[0].IpConfigurations.Count -gt 1) {
            #Version 1.08 - Ensure that if there is more than one then have it be a string rather than array
            $nic_1_id_ip2 = ($NICInfo[0].IpConfigurations | Where { $_.Primary -eq $False }).PrivateIpAddress -join "`n"
        } Else {
            $nic_1_id_ip2 = "none"
        }

        #If there is at least one NIC
        if ($vm.NetworkProfile.NetworkInterfaces.Count -gt 1) {
            ###Store data for NIC 2###
            $nic_2_Name = $NICInfo[1].Name
            $nic_2_Tag = $NICInfo[1].Tag."App-name"
            #Try puting the Accelerated Network setting below this note
            $nic_2_AcceleratedNetworking = $NICInfo[1].EnableAcceleratedNetworking
            #Store the primary IP as IP1
            $nic_2_id_ip1 = ($NICInfo[1].IpConfigurations | Where { $_.Primary -eq $True }).PrivateIpAddress
            $nic_2_id_subnet = $NICInfo[1].IpConfigurations.Subnet.ID | Split-Path -Leaf | Select -First 1
            #If there are multiple IPs on a NIC
            If ($NICInfo[1].IpConfigurations.Count -gt 1) {
                #Version 1.08 - Ensure that if there is more than one then have it be a string rather than array
                $nic_2_id_ip2 = ($NICInfo[1].IpConfigurations | Where { $_.Primary -eq $False }).PrivateIpAddress -join "`n"
            } Else {
                $nic_2_id_ip2 = "none"
            }
        } Else {
            Clear-Variable nic_2_Name, nic_2_Tag, nic_2_AcceleratedNetworking, nic_2_id_ip1, nic_2_id_subnet, nic_2_id_ip2 -ErrorAction SilentlyContinue
        }
        
        ###Store data for NIC 3###
        if ($vm.NetworkProfile.NetworkInterfaces.Count -gt 2) {
            $nic_3_Name = $NICInfo[2].Name
            $nic_3_Tag = $NICInfo[2].Tag."App-name"
            #Try puting the Accelerated Network setting below this note
            $nic_3_AcceleratedNetworking = $NICInfo[2].EnableAcceleratedNetworking
            #Store the primary IP as IP1
            $nic_3_id_ip1 = ($NICInfo[2].IpConfigurations | Where { $_.Primary -eq $True }).PrivateIpAddress
            $nic_3_id_subnet = $NICInfo[2].IpConfigurations.Subnet.ID | Split-Path -Leaf | Select -First 1
            #If there are multiple IPs on a NIC
            If ($NICInfo[2].IpConfigurations.Count -gt 1) {
                #Version 1.08 - Ensure that if there is more than one then have it be a string rather than array
                $nic_3_id_ip2 = ($NICInfo[2].IpConfigurations | Where { $_.Primary -eq $False }).PrivateIpAddress -join "`n"
            } Else {
                $nic_3_id_ip2 = "none"
            }
        } Else {
            Clear-Variable nic_3_Name, nic_3_Tag, nic_3_AcceleratedNetworking, nic_3_id_ip1, nic_3_id_subnet, nic_3_id_ip2 -ErrorAction SilentlyContinue
        }

        #Version 1.02 - Added three additional fields, SecurityType, SecureBoot and vTPMStatus
        $SecurityType = $vm.SecurityProfile.SecurityType
        $SecureBoot = $vm.SecurityProfile.UefiSettings.SecureBootEnabled
        $VTPMStatus = $vm.SecurityProfile.UefiSettings.vTpmEnabled

        $AzureVMInventoryList.add((New-Object "psobject" -Property @{
            "VMName"=$vm_name;"ResourceGroupName"=$vm_rg;"VMStatus"=$vmstatus;"Location"=$vm.Location;
            "App-name"=$application;"Function"=$function;"Impact"=$impact;"Mission"=$mission;
            "LicenseType"=$vm.LicenseType;"VMSize"=$vmsize;"ManagedDisk"=$managed;"VM_Memory"=$vmdata.MemoryInMB;
            "VM_Cores"=$vmdata.NumberOfCores;"VMAgentVersion"=$agent;"VM_Agent_Status"=$agentstatus;"OSType"=$vm.StorageProfile.OsDisk.OsType;
            "OSDisk"=$vm.StorageProfile.OsDisk.name;"OSDiskSize"=$vm.StorageProfile.OsDisk.DiskSizeGb;"BootDiag"=$bd;"BootDiagSA"=$sa;
            "Shutdown"=$vm.Tags.Shutdown;"Startup"=$vm.Tags.Startup;"Role"=$vmrole;"Role_Assignment"=$vmDisplayName;
            "VM_Generation"=$generation;"Nic_1_Name"=$nic_1_name;"Nic_1_Tag"=$nic_1_tag;"Nic_1_AcceleratedNetworking"=$nic_1_AcceleratedNetworking;
            "Nic_1_Subnet"=$nic_1_id_subnet;"Nic_1_IP1"=$nic_1_id_ip1;"Nic_1_IP2"=$nic_1_id_ip2;"Nic_2_Name"=$nic_2_name;"Nic_2_Tag"=$nic_2_tag;
            "Nic_2_AcceleratedNetworking"=$nic_2_AcceleratedNetworking;"Nic_2_Subnet"=$nic_2_id_subnet;"Nic_2_IP1"=$nic_2_id_ip1;
            "Nic_2_IP2"=$nic_2_id_ip2;"Nic_3_Name"=$nic_3_name;"Nic_3_Tag"=$nic_3_tag;"Nic_3_AcceleratedNetworking"=$nic_3_AcceleratedNetworking;
            "Nic_3_Subnet"=$nic_3_id_subnet;"Nic_3_IP1"=$nic_3_id_ip1;"Nic_3_IP2"=$nic_3_id_ip2;"OS_name"=$OS_Name;
            "OS_Version"=$OS_Version;"SubScriptionName"=$SubscriptionShortName;"SecurityType"=$SecurityType;"SecureBoot"=$SecureBoot;
            "vTPMStatus"=$VTPMStatus}))
    } #End of foreach VM
} #End of foreach Subscription
Write-Host "`nWriting data to Excel" -ForegroundColor Cyan

#Move the old reports into the SavedFiles path
Get-ChildItem -Path $ReportPath -Filter "$ShortReportName*" | ForEach {
    Move-Item -Path $_.FullName "$ReportPath\SavedFiles" -Force
}

#Set the data headers in the correct order
$FinalOutput = $AzureVMInventoryList | Select $ReportColumnHeaders | Sort "VMName"

#Loop through each Azure subscription and 
ForEach ($CurrSubscription in $Subscriptions) {
#    #Get the subscription short name
    $SubscriptionShortName = $CurrSubscription.Name.split('-')[1]
    #Report data for just this subscription, added FreezePane option to the Export-Excel
    $FinalOutput | Where { $_."SubscriptionName" -Like "*$SubscriptionShortName" } | Export-Excel -WorksheetName $SubscriptionShortName -Path "ReportPath:\$ReportFileName" -AutoSize -Append -FreezePane 2, 2
}

#Build the special sheets
#All_VMs
$FinalOutput | Export-Excel -WorksheetName "All_VMs" -Path "ReportPath:\$ReportFileName" -AutoSize -Append -FreezePane 2, 2
#Not_Running
$FinalOutput | Where { $_.VMstatus -notlike "running" } | Export-Excel -WorksheetName "Not_Running" -Path "ReportPath:\$ReportFileName" -AutoSize -Append -FreezePane 2, 2
#VMs_NoApplicationTag, updated to put "None" in the Application cell
$FinalOutput | Where { $_."App-name" -eq "None" } | Export-Excel -WorksheetName "NO_App-name" -Path "ReportPath:\$ReportFileName" -AutoSize -Append -FreezePane 2, 2
#VMs_NoBootDiag
$FinalOutput | Where { $_.BootDiag -eq "none" } | Export-Excel -WorksheetName "NO_BootDiag" -Path "ReportPath:\$ReportFileName" -AutoSize -Append -FreezePane 2, 2
#IL4_VMs 
$FinalOutput | Where { $_.Impact -eq "IL4" } | Export-Excel -WorksheetName "IL4_VMs" -Path "ReportPath:\$ReportFileName" -AutoSize -Append -FreezePane 2, 2
#IL5_VMs 
$FinalOutput | Where { $_.Impact -eq "IL5" } | Export-Excel -WorksheetName "IL5_VMs" -Path "ReportPath:\$ReportFileName" -AutoSize -Append -FreezePane 2, 2


#Set auto filter on all worksheets
#Open the spreadsheet that was created
$ExcelPkg = Open-ExcelPackage -Path "ReportPath:\$ReportFileName"
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

Write-Host "`nScript Complete.  Report written to $ReportPath\$ReportFileName" -ForegroundColor Green

#Stop the transcript
Stop-Transcript
# SIG # Begin signature block
# MIIL6gYJKoZIhvcNAQcCoIIL2zCCC9cCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAX2EkXu5gvq1yW
# KO2f7YlxWBfZqhQIgFT9bOnuRqbFm6CCCS0wggRsMIIDVKADAgECAgMSNG8wDQYJ
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
# SIb3DQEJBDEiBCCNi58jinGZKHd/uO5/HfLKxE8VrPB2a3gJB/rR72dVpjANBgkq
# hkiG9w0BAQEFAASCAQAT1zY7BHCAa0xbPlxlE0BzJTzFg7h3gM+HY3qmdOb8Rjt5
# L62BVR5ry7Z9mHi19KUBI9BvxOkGqldM25GxSQR4yqgTIcYFeFkbwSUfLY/BRh0w
# t/mfIO/oNPe+nLh4jsl4lsAlUQoNKhV+vtRinG//NuOJXSRv20jUND0c9ku2hAAb
# BfWefKr9GQdJ4uvcYjeb60Sewj5UITon5WwjX3eR4koaAYMbkqG/f4vQP/sAGEcW
# hSrBh9WSo6kqKhUQtzyv48prd27BTzOhSSfTwxOp3af8ovNR61CoFdi72qVIpEoM
# gQgA76qWViUEBCILvDgWa84tVY93PFMo4W1cwN8t
# SIG # End signature block
