#Purpose: Launch the Get_All_ESX_Inventory script

#ChangeLog
#Version 1.00 - 08/15/24 - MDR - Initial version
#Version 1.01 - 09/12/24 - MDR - Adding in the svc_vrapsh password argument
#Version 1.02 - 10/03/24 - MDR - Converting passwords into Base64 and exporting them to a .csv

#Version 1.02 - If C:\Temp doesn't exist then create it
If (Test-Path "C:\Temp") {
    New-Item "C:\Temp" -ErrorAction SilentlyContinue -ItemType Directory | Out-Null
}

#Version 1.01 - Create an ouput file for commands being executed
Start-Transcript C:\Temp\ESX_Inventory_Launcher.txt

[string]$vCenterPassword = $args[0]
#Version 1.01 - Get svc_vrapsh password argument
[string]$VRAPassword = $args[1]

#Version 1.02 - Define where the parameters will be temporarily stored
$ParameterPath = "C:\Temp\ESXInventoryParams.csv"

#Version 1.02 - Convert passwords into Base64
$vCenterPasswordBase64 = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($vCenterPassword))
$VRAPasswordBase64 = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($VRAPassword))

#Version 1.02 - Export the Base64 passwords to a .csv
[PSCustomObject]@{"vCenterPassword"=$vCenterPasswordBase64;"VRAPassword"=$VRAPasswordBase64} | Export-Csv $ParameterPath

# Path to the PowerShell 7 executable
$PS7Path = "C:\Program Files\PowerShell\7\pwsh.exe"

# Path to your PowerShell 7 script
$ScriptPath = "C:\DLA-Failsafe\vRA\MikeR\Get_All_ESX_Inventory_v1.10.ps1"

# Use Start-Process to run the script with PowerShell 7
Start-Process -FilePath $PS7Path -ArgumentList "$ScriptPath -ParameterPath $ParameterPath -RunFromOrchestrator True" -Wait

#Version 1.01 - Stop the transcript
Stop-Transcript
# SIG # Begin signature block
# MIIL6gYJKoZIhvcNAQcCoIIL2zCCC9cCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCtLOpZYNudOn8p
# ayUBwBOzo5D5zsRxtb0iqz94Ya/b+qCCCS0wggRsMIIDVKADAgECAgMSNG8wDQYJ
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
# SIb3DQEJBDEiBCBrHXrf6bxVYHHimK8CmVdftOUtMPzHOQpNzMxjoivk0jANBgkq
# hkiG9w0BAQEFAASCAQAEZWUUcRdiEVnHMw2SmgAwYllb2rDSxzoZ/VR0s8ZGL+t1
# gLQyQqllSJ0wlua4V3z1guwAfOYBYTtYn5KvjrcewzWwK0lWzShqdxsA4JK1+gJx
# GDAOkktgl9601KaJPAE5HnUdfrYZkqrcrHaoTG/x0dCqrH6bLiRBYobe0PKDLvgC
# YVp6xA9FJJUKBttdBhMakKNffmjbxATWJoSZ7WyoSRJTsEOdNHPgRRtX20O5fTQ/
# 3/Nv6So2dcTKGj3Wpe4BeeTsui42xZDGJ/zk7d7E9hnUOt8Fk5rGUS/ORcyC08ne
# XAsGzOJAjC5YtiyFXFs7bIPDmBnPNqxfEo/M09dG
# SIG # End signature block
