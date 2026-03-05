#PURPOSE: This script will collect the AWS VM inventory

#CHANGELOG
#Version 1.00 - 09/17/24 - MDR - Initial version
#Version 1.01 - 10/25/24 - MDR - Added "-FreezePane 2, 2" to all Export-Excel lines
#Version 1.02 - 10/25/24 - MDR - If the ReportPath PSDrive exists then remove it before recreating it
#Version 1.03 - 10/25/24 - MDR - Make this script compatible with running from Orchestrator
#Version 1.04 - 10/28/24 - MDR - If the script fails to connect to AWS then exit the script
#Version 1.05 - 05/29/25 - MDR - Added a line to remove the drive mapping at the end of the script
#Version 1.06 - 10/14/25 - MDR - Store report to local temp folder before moving it to the network share
#Version 1.07 - 02/23/26 - MDR - Prevent the script from attempting to write report data if an error occurred earlier preventing data from being collected

Param ( $ParameterPath, $RunFromOrchestrator )

#Version 1.03 - Only import parameters from file if run from Orchestrator
If ($RunFromOrchestrator -eq "True") {
    #Import Base64 passwords from CSV
    $Base64Passwords = Import-CSV $ParameterPath

    #Delete the CSV file since it is no longer needed and for security reasons
    Remove-Item $ParameterPath | Out-Null

    #Store passwords to temp variables
    $VRAPasswordBase64 = $Base64Passwords.VRAPassword

    #Decode passwords from Base64 to plain text
    $VRAPassword = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($VRAPasswordBase64))
}

#Configure variables
$TodaysDate = Get-Date -Format "MMddyyyy"
$ReportPath = "\\orgaze.dir.ad.dla.mil\DCC\VirtualTeam\Reports"
$ReportFileName = "Master_AWS_VMInventory_$TodaysDate.xlsx"
$ShortReportName = "Master_AWS_VMInventory"
$AWSInventoryData = New-Object System.Collections.Generic.List[System.Object]

#Report column headers
$ReportColumnHeaders = "AWS_Instance","DNS_Name","Tag_Name","Tag_Description","Tag_OS","Tag_Backup","AWS_Type","Status","Image","Platform","IPAddress","Security_Group","CPU_Cores","CPU_Threads","Availabity_Zone","Tag_Team",
                       "Tag_ECR","Tag_Builder","Tag_AMI","Tag_ConfigTest"

#Version 1.06 - Ensure Temp folder exists
If (!(Test-Path "C:\Temp")) {
    New-Item "C:\Temp" -ItemType Directory | Out-Null
}

Start-Transcript C:\Temp\Get_All_AWS_VMInventory_$TodaysDate.txt

Clear-Host

#If a ReportPath drive mapping already exists then remove it
Get-PSDrive -Name ReportPath -ErrorAction SilentlyContinue | Remove-PSDrive

#If this is run from Orchestrator then generate a credential to perform the PSDrive mapping
If ($RunFromOrchestrator -eq "True") {
    #Create a credential for svc_vrapsh
    $VRACred = New-Object System.Management.Automation.PSCredential -ArgumentList @("DIR\svc_vrapsh",(ConvertTo-SecureString -String $VRAPassword -AsPlainText -Force))
    #Map a drive to the ReportPath
    New-PSDrive -Name "ReportPath" -PSProvider FileSystem -Root $ReportPath -Credential $VRACred | Out-Null
} Else { #If not run from Orchestrator then just map the PSDrive with current credentials
    #Version 1.02 - If the ReportPath PSDrive exists then remove it before recreating it
    If (Test-Path ReportPath:) {
        Set-Location C:
        Remove-PSDrive ReportPath
    }
    #Map a drive to the ReportPath
    New-PSDrive -Name "ReportPath" -PSProvider FileSystem -Root $ReportPath -ErrorAction SilentlyContinue | Out-Null
}

If (!(Get-PSDrive -Name "ReportPath")) {
    Write-Host "Failed to connect to ReportPath.  Exiting" -ForegroundColor Red
    Break
}

#If for some reason an Access Denied error happens here then exit the script
Try {
    #If this report exists already then delete it
    If (Test-Path "ReportPath:\$ReportFileName") {
        Remove-Item "ReportPath:\$ReportFileName" -ErrorAction Stop
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

#If ImportExcel is not found then prompt for folder where it is located
If (!$CheckForImportExcel) {
    Write-Host "The ImportExcel module is required for this script to run" -ForegroundColor Red
    Write-Host "`nA copy of this module is located in \\orgaze.dir.ad.dla.mil\J6_INFO_OPS\J64\J64C\WinAdmin\VulnMgt\Software\ImportExcel"
    Write-Host "`nPlace a copy of this module in C:\Program Files\WindowsPowerShell\Modules"
    Break
}

#Check to see if ImportExcel is installed
$CheckForAWSTools = Get-Command Get-EC2Instance -ErrorAction SilentlyContinue

#If ImportExcel is not found then prompt for folder where it is located
If (!$CheckForAWSTools) {
    Write-Host "The AWS Tools module is required for this script to run" -ForegroundColor Red
    Write-Host "`nA copy of this module is located in \\orgaze.dir.ad.dla.mil\DCC\VirtualTeam\Scripts\Modules\AWS Admin Tools"
    Write-Host "`nFollow the instructions in 'AWS Tools install instructions.txt'"
    Break
}

#Store the list of regions that need data collected for it
$AWSRegions =  "us-gov-west-1", "us-gov-east-1"

#Loop through all regions
ForEach ($Region in $AWSRegions) {
    #Output Region name
    Write-Host "`nScanning $Region"

    #Set the default region
    Set-DefaultAWSRegion -Region $Region

    #Version 1.07 - If an error occurs getting instances then report the error and exit
    Try {
        #Get a list of all AWS instances
        $AllInstances = Get-EC2Instance -ErrorVariable AWSError | Select -ExpandProperty Instances | Sort
    } Catch {
        Write-Host "Error connecting to AWS: $AWSError" -ForegroundColor Red
        Break
    }

    #Populate variables for progress bar
    $TotalInstances = $AllInstances.Count
    $CurrentInstanceNum = 1

    ForEach ($EC2Instance in $AllInstances) {
        #If this is run from Orchestrator then output the Instance name to get written to the transcript file
        If ($RunFromOrchestrator -eq "True") {
            Write-Host "$(($EC2Instance.Tags | Where { $_.Key -eq "Name" }).Value)"
        } Else { #If not from Orchestrator then write a progress bar
            Write-Progress -Activity "Getting instance data" -Status "$CurrentInstanceNum of $TotalInstances" -PercentComplete ($CurrentInstanceNum / $TotalInstances * 100)
        }

        #Collect data about the current instance
        $InstanceData = Get-EC2Instance -InstanceId $EC2Instance.InstanceId
        #Get the instance ID
        $InstanceID = $InstanceData.Instances.InstanceID
        #Get the instance type
        $InstanceType = ($EC2Instance.InstanceType).Value
        #Get the instance status (ie running)
        $InstanceStatus = ((Get-EC2InstanceStatus -InstanceId $instanceID).InstanceState.Name.Value)
        #If the instance isn't "running" then set it to "not running"
        If ($InstanceStatus -ne "running") {
            $InstanceStatus = "Not Running"
        }
        $InstanceImageID = $EC2Instance.ImageID
        $InstancePlatform = $EC2Instance.Platform
        $IPaddress = $EC2Instance.PrivateIpAddress
        If ($IPaddress -ne $null) {
            $DNSName = (Resolve-DNSName -Name $IPaddress -DnsOnly -ErrorAction SilentlyContinue).NameHost
        }
        #If there is more than one DNS entry then find one that starts with AW and is more than 3 charaters long.  This prevents DCs from just reporting back the domain name rather than the host name
        If ($DNSName.Count -gt 1) {
            $DNSName = $DNSName | Where { $_ -like "AW*" -and $_.Split(".")[0].Length -gt 3 }
            #If there is more than one match to this then just take the first one
            If ($DNSName.Count -gt 1) {
                $DNSName = $DNSName[0]
            }
        }
        #Get Tag info
        $Tag_Name = ($EC2Instance.Tags | Where {$_.Key -eq "Name"}).Value
        $Tag_Desc = ($EC2Instance.Tags | Where {$_.Key -eq "Description"}).Value
        $Tag_OS = ($EC2Instance.Tags | Where {$_.Key -eq "OS"}).Value
        $Tag_Backup = ($EC2Instance.Tags | Where {$_.Key -eq "Backup"}).Value
        $Tag_Team = ($EC2Instance.Tags | Where {$_.Key -eq "Team"}).Value
        $Tag_ECR = ($EC2Instance.Tags | Where {$_.Key -eq "ChangeRequest"}).Value
        $Tag_Builder = ($EC2Instance.Tags | Where {$_.Key -eq "Builder"}).Value
        $Tag_AMI = ($EC2Instance.Tags | Where {$_.Key -eq "AMI"}).Value
        $Tag_ConfigTest = ($EC2Instance.Tags | Where {$_.Key -eq "configtest"}).Value
        #Get group names
        $SecurityGroups = ($EC2Instance.SecurityGroups).GroupName
        #Get CPU data
        $CPUOptions = ($EC2Instance).CPUOptions
        #Get availability zone
        $AvailabilityZone = $EC2Instance.Placement.AvailabilityZone

        #Store data from the instance
        $AWSInventoryData.add((New-Object "psobject" -Property @{"AWS_Instance"=$InstanceID;
        "DNS_Name"=$DNSName;"Tag_Name"=$Tag_Name;"Tag_Description"=$Tag_Desc;"Tag_OS"=$Tag_OS;"Tag_Backup"=$Tag_Backup;"AWS_Type"=$InstanceType;"Status"=$InstanceStatus;"Image"=$InstanceImageID;
        "Platform"=$InstancePlatform;"IPAddress"=$IPaddress;"Security_Group"=$SecurityGroups;"CPU_Cores"=$CPUOptions.CoreCount;"CPU_Threads"=$CPUOptions.ThreadsPerCore;"Availabity_Zone"=$AvailabilityZone;
        "Tag_Team"=$Tag_Team;"Tag_ECR"=$Tag_ECR;"Tag_Builder"=$Tag_Builder;"Tag_AMI"=$Tag_AMI;"Tag_ConfigTest"=$Tag_ConfigTest}))

        #Increment the current instance number for the progress bar
        $CurrentInstanceNum++
    }

    #Close the progress bar
    Write-Progress -Activity "Getting instance data" -Completed
}

#Version 1.07 - Only perform report operations if there is data to report
If ($FinalOutput -eq $null) {
    Write-Host "`nNo data to report" -ForegroundColor Red
    Break
}

Write-Host "`nWriting data to Excel" -ForegroundColor Cyan

#Move the old reports into the SavedFiles path
Get-ChildItem -Path $ReportPath -Filter "$ShortReportName*" | ForEach {
    Move-Item -Path $_.FullName "$ReportPath\SavedFiles" -Force
}

#Adding Set-Location to ReportPath: for some reason helps ensure no errors when exporting data
Set-Location ReportPath:

#Set the data headers in the correct order
$FinalOutput = $AWSInventoryData | Select $ReportColumnHeaders

#All VM Hosts
$FinalOutput | Export-Excel -WorksheetName "All_VMs" -Path "C:\Temp\$ReportFileName" -AutoSize -Append -FreezePane 2, 2
#Get instances in east
$FinalOutput | Where { $_.Availabity_Zone -like "*east*" } | Export-Excel -WorksheetName "EAST" -Path "C:\Temp\$ReportFileName" -AutoSize -Append -FreezePane 2, 2
#Get instances in wast
$FinalOutput | Where { $_.Availabity_Zone -like "*west*" } | Export-Excel -WorksheetName "WEST" -Path "C:\Temp\$ReportFileName" -AutoSize -Append -FreezePane 2, 2

#Version 1.06 - Only do auto filter if not running from the PSH host
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

#Version 1.06 - Move the report to the report folder
Move-Item -Path "C:\Temp\$ReportFileName" -Destination "ReportPath:\"

#Version 1.05 - Remove ReportPath
Set-Location c:
Remove-PSDrive -Name "ReportPath"

Write-Host "`nScript Complete.  Report written to $ReportPath\$ReportFileName" -ForegroundColor Green

#Stop the transcript
Stop-Transcript
# SIG # Begin signature block
# MIIL6gYJKoZIhvcNAQcCoIIL2zCCC9cCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCCg2xWj16UhuHK
# o6fCZ4UJMTKqcDnV0FVuNgMOcH1OfaCCCS0wggRsMIIDVKADAgECAgMSNG8wDQYJ
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
# SIb3DQEJBDEiBCCtdXas5jSyEVEcva4Syv6gOAShLttu49Ky4V9M6vZo7jANBgkq
# hkiG9w0BAQEFAASCAQBe1eDNkkicgc9fkdxARzC2g2qnqOCOnP6rzDJq9bEAqaYY
# w+T7OxxRpyna9DkcyvNYfTJ9iKCwcRBMugBYCtoCCgrg2/JZ5EsWpShbmLAlxBLV
# kHfZCMoThc+sYbsBqqfhU66XJ9UAWUVJHrArLiOF5OhV3jgCwn+1d5dpaz83eC3p
# DwE2pyQwyQQfBANl9GZNMFHypiUSvZ5XwJcSZSlDwlnfz2WIzNGnX35FaZFbbqjh
# VsZswmy5HBbl37ZBBMp2cY+IshvY0WvGKCkrqfBxGUycBb1leZLAEf4XnocM17dc
# baB6UCX50E40R+O+jK8o0zsE1Drsxrb9++VSflsf
# SIG # End signature block
