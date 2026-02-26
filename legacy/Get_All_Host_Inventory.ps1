#JN 12/11/2019 This version of the srcipt was copied from vs346 and modified to run on the new Windows 2016 jump server
#   only change was from D Drive to E Drive (D: > E:) in multiple locations, working out of Task Scheduler for limited vCenters

# This script will run against all the vCenters defined by $vcs, this must be run using a DIR M account
#  The following line creates a output of the commands executed during the script processing
start-transcript -path E:\Script_Output\ESX_HostInventoryFiles\Get_All_Host_Inventory.txt

#Declare some variables
$date = (Get-Date).ToString("HH:mm:ss MM-dd-yy")
write-host "Start of Get_All_Host_Inventory.ps1..." $date

#List of all vCenters, HAWAII vCenter hnl1s-ph1005v not working 
   $vcs =  "e:\listofVcenterServers.csv"


#Function to do some file cleanup from previous runs
Function TestForFiles{
	    $Sourcefile = "E:\ESX_HostInventory_$vc.csv"
	    $Destinationfile = "E:\ESX_HostInventory_$vc.csv"

	If ((Test-Path $Destinationfile) -eq $false)  {
		write-host "No destination file exists"
		# Previous files does not exist so move the existing one prior to creating a new
		# First test if it even exists
		if ((Test-Path $Sourcefile) -eq $false) {
			Write-host "No file exist to move"
		}
		else {
			write-host "A sourcefile exist, move it"
			Move-Item -Path $Sourcefile -Destination $Destinationfile
		}
  	}
  	else  {
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
  	} #End of if test for previous files 
} #End of Function

#Main portion of the script
Write-host -ForeGroundColor Cyan "`n`tThis script will get the inventory of the ESX Hosts on each vCenter"
foreach ($vc in $vcs) {
	write-host "Running Get_HostInventory on $vc"
	#Clean up from previous execution of this script
	TestForFiles $vc
	if ($vc -eq "vcentername1") {
        $creds = Get-VICredentialStoreItem -Host vcentername1
		connect-viserver -Server $vc -User $creds.User -Password $creds.password
    }
	elseif ($vc -eq "vcentername2") {
        $creds = Get-VICredentialStoreItem -Host vcentername2 
		connect-viserver -Server $vc -User $creds.User -Password $creds.password
    }
    elseif ($vc -eq "vcentername3") {
		$creds = Get-VICredentialStoreItem -Host 
	    connect-viserver -server $vc
	}
	else {
		write-host -ForeGroundColor Yello "Invalid vCenter being proessed..."
        exit
	} #End of IF/ELSEIF/ELSE tests
    
	E:\Scripts\Get_HostInventory.ps1 > $null
    
    #Rename output file to include the hostname of vCenter
    move E:\ESX_HostInventory.csv E:\ESX_HostInventory_$vc.csv
    
	Disconnect-viServer -Server * -Confirm:$false
} #End of foreach loop

$date = (Get-Date).ToString("HH:mm:ss MM-dd-yy")
write-host "End of Get_All_Host_Inventory.ps1..." $date

#Stops the output of commands
stop-transcript