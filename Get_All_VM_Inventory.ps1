
# This script will run against all the vCenters defined by $vcs, this must be run using a domain account
#  The following line creates a output of the commands executed during the script processing
start-transcript -path E:\Script_Output\VM_InventoryFiles\Get_All_VM_Inventory.txt
#start-transcript -path E:\Script_Output\VM_InventoryFiles\Get_All_VM_Inventory_RunOnPSHost.txt

#Declare some variables
    #Thia is to show the actual date and script that is running
    $date = (Get-Date).ToString("HH:mm:ss MM-dd-yy")
    write-host "Start of Get_All_VM_Inventory.ps1..." $date
#    write-host "Start of E:Scripts\PRODUCTION\Get_All_VM_Inventory_RunOnPSHost.ps1..." $date

#List of all vCenters, HAWAII vCenter hnl1s-ph1005v not working 
   $vcs =  "nameofvcenter1", "nameofvcenter2", "nameofvcenter3", "nameofvcenter4", "nameofvcenter5", "nameofvcenter6", "nameofvcenter7", "nameofvcenter8"

  

#Function to do some file cleanup from previous runs
Function TestForFiles{
	$Sourcefile = "E:\Script_Output\VM_InventoryFiles\VMinventory_$vc.csv"
	$Destinationfile = "E:\Script_Output\VM_InventoryFiles\PriorVMinventoryFiles\VMinventory_$vc.csv"
	
	If ((Test-Path $Destinationfile) -eq $false) {
		write-host "No destination file exists"
		# Previous files does not exist so move the existing one prior to creating a new
		# First test if it even exists
		if ((Test-Path $Sourcefile) -eq $false)	{
			Write-host "No file exist to move"
		}
		else {
			write-host "A sourcefile exist, move it.."
			Move-Item -Path $Sourcefile -Destination $Destinationfile
		}
  	}
  	else {
		# Previous file in the Destination exist, delete it then move the one in the source directory
		write-host "Destination file exists, remove it"
		Remove-Item $Destinationfile
		# First test if a sourcefile exists
		if ((Test-Path $Sourcefile) -eq $false) {
			Write-host "No previous file exist"
		}
		else {
			write-host "Source file exists, move it"
			Move-Item -Path $Sourcefile -Destination $Destinationfile
		}
  	}
} #End of Function

#Main portion of the script
    Write-host "This script will get the inventory of the VM's on each vCenter"
foreach ($vc in $vcs) {
	write-host "Running Get_VM_Inventory on $vc"
	TestForFiles
    sleep 10 
	if ($vc -eq "nameofvcenter") {
		$creds = Get-VICredentialStoreItem -Host nameofvcenter -User domain\user -File C:\Users\AppData\Roaming\VMware\credstore\vicredentials.xml
        write-host "Credentials returned..." $creds.Host $creds.User $creds.Password
		connect-viserver -Server $vc -User $creds.User -Password $creds.password
    }
    elseif ($vc -eq "nameofvcenter") {
		$creds = Get-VICredentialStoreItem -Host nameofvcenter -User "administrator@ssodomain" -File C:\Users\AppData\Roaming\VMware\credstore\vicredentials.xml
        write-host "Credentials returned..." $creds.Host $creds.User $creds.Password
	    connect-viserver -server $vc -user $creds.User -Password $creds.password
	}
    elseif ($vc -eq "nameofvcenter") {
		$creds = Get-VICredentialStoreItem -Host nameofvcenter -User domain\user -File C:\Users\AppData\Roaming\VMware\credstore\vicredentials.xml
        write-host "Credentials returned..." $creds.Host $creds.User $creds.Password
	    connect-viserver -server $vc -user $creds.User -Password $creds.password
	}
    elseif ($vc -eq "nameofvcenter") {
		$creds = Get-VICredentialStoreItem -Host nameofvcenter -User domain\user -File C:\Users\AppData\Roaming\VMware\credstore\vicredentials.xml
        write-host "Credentials returned..." $creds.Host $creds.User $creds.Password
	    connect-viserver -server $vc -User $creds.User -Password $creds.Password
	}
    elseif ($vc -eq "nameofvcenter") {
		$creds = Get-VICredentialStoreItem -Host nameofvcenter -User domain\user -File C:\Users\AppData\Roaming\VMware\credstore\vicredentials.xml
        write-host "Credentials returned..." $creds.Host $creds.User $creds.Password
	    connect-viserver -server $vc -User $creds.User -Password $creds.Password
	}
    elseif ($vc -eq "nameofvcenter") {
		$creds = Get-VICredentialStoreItem -Host nameofvcenter -User domain\user -File C:\Users\AppData\Roaming\VMware\credstore\vicredentials.xml
        write-host "Credentials returned..." $creds.Host $creds.User $creds.Password
		connect-viserver -server $vc -User $creds.User -Password $creds.Password
	}
    elseif ($vc -eq "nameofvcenter") {
		$creds = Get-VICredentialStoreItem -Host nameofvcenter -User domain\user -File C:\Users\AppData\Roaming\VMware\credstore\vicredentials.xml
        write-host "Credentials returned..." $creds.Host $creds.User $creds.Password
		connect-viserver -server $vc -User $creds.User -Password $creds.Password
	}
    elseif ($vc -eq "nameofvcenter") {
		$creds = Get-VICredentialStoreItem -Host nameofvcenter -User "administrator@ssoadmin" -File C:\Users\AppData\Roaming\VMware\credstore\vicredentials.xml
        write-host "Credentials returned..." $creds.Host $creds.User $creds.Password
	    connect-viserver -server $vc -user $creds.User -Password $creds.password
	}
    elseif ($vc -eq "nameofvcenter") {
		$creds = Get-VICredentialStoreItem -Host nameofvcenter -User domain\user -File C:\Users\AppData\Roaming\VMware\credstore\vicredentials.xml
        write-host "Credentials returned..." $creds.Host $creds.User $creds.Password
		connect-viserver -server $vc -User $creds.User -Password $creds.Password
	}
	else {
		write-host -ForeGroundColor Yello "Invalid vCenter being proessed..."
        exit
	} #End of IF/ELSEIF/ELSE tests

	E:\Scripts\Get_VM_Inventory.ps1 > $null
    #E:\Scripts\PRODUCTION\Get_VM_Inventory_RunOnPSHost.ps1

	#Previous line creates a generic output file, need to rename with vcenter being processed
	move E:\Script_Output\VM_InventoryFiles\VMinventory.csv E:\Script_Output\VM_InventoryFiles\VMinventory_$vc.csv
	
	Disconnect-viServer -Server * -Confirm:$false
}  #End of foreach loop

$date = (Get-Date).ToString("HH:mm:ss MM-dd-yy")
write-host "End of E:\Scripts\Get_All_VM_Inventory.ps1..." $date
#write-host "End of E:\Scripts\PRODUCTION\Get_All_VM_Inventory_RunOnPSHost.ps1..." $date

#Stops the output of commands
stop-transcript