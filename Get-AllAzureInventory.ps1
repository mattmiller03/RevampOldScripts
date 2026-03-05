#Requires -Modules Az.Accounts, Az.Compute, Az.Network, Az.Resources, ImportExcel

<#
.SYNOPSIS
    Collects Azure VM inventory across all subscriptions and produces a multi-tab Excel workbook.

.DESCRIPTION
    Iterates through each Azure subscription listed in config/azure.json, connects using
    DPAPI-encrypted service principal credentials (created by Initialize-AzureCredentials.ps1),
    collects VM inventory, and produces a single Excel workbook with multiple tabs:

      - Search       : VBA-powered search UI (searches All_VMs table)
      - All_VMs      : Combined VM inventory from all subscriptions
      - Not_Running  : VMs not in 'running' state
      - NO_App-name  : VMs missing the App-name tag
      - NO_BootDiag  : VMs without boot diagnostics configured
      - IL4_VMs      : VMs with Impact-level tag = IL4
      - IL5_VMs      : VMs with Impact-level tag = IL5
      - <Subscription> : One tab per subscription

.PARAMETER ConfigFile
    Path to the JSON configuration file. Defaults to config/azure.json.

.PARAMETER OutputDir
    Directory where the inventory workbook is written. Created if it does not exist.

.PARAMETER ArchiveDir
    Directory where the previous workbook is moved before a new run.

.PARAMETER TranscriptDir
    Directory for transcript log files.

.EXAMPLE
    .\Get-AllAzureInventory.ps1
    Runs with default paths relative to the script directory.

.EXAMPLE
    .\Get-AllAzureInventory.ps1 -OutputDir "D:\Reports\AzureInventory" -Verbose
    Runs with a custom output directory and verbose logging.
#>

[CmdletBinding()]
param(
    [Parameter()]
    [string]$ConfigFile = (Join-Path $PSScriptRoot 'config\azure.json'),

    [Parameter()]
    [string]$OutputDir = (Join-Path $PSScriptRoot 'Output\AzureInventory'),

    [Parameter()]
    [string]$ArchiveDir = (Join-Path $PSScriptRoot 'Output\AzureInventory\Archive'),

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
$transcriptPath = Join-Path $TranscriptDir "Get-AllAzureInventory_$(Get-Date -Format 'yyyy-MM-dd_HHmmss').log"
Start-Transcript -Path $transcriptPath

$startTime = Get-Date
Write-Host "Start of Get-AllAzureInventory.ps1 at $($startTime.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor Cyan

# Load configuration
if (-not (Test-Path -Path $ConfigFile)) {
    Write-Error "Configuration file not found: $ConfigFile"
    Stop-Transcript
    return
}

$config = Get-Content -Path $ConfigFile -Raw | ConvertFrom-Json
$credDir = Join-Path $PSScriptRoot $config.CredentialDir
$totalSubs = ($config.Tenants.Subscriptions | Measure-Object).Count
Write-Host "Loaded $($config.Tenants.Count) tenant(s) / $totalSubs subscription(s) from config." -ForegroundColor Cyan
Write-Host "Environment: $($config.Environment)" -ForegroundColor Cyan

# Output file paths
$reportFile = Join-Path $OutputDir 'AzureInventory_All.xlsm'
$archiveFile = Join-Path $ArchiveDir 'AzureInventory_All.xlsm'

# Archive previous report
Backup-PreviousReport -SourcePath $reportFile -ArchivePath $archiveFile

#endregion Initialization

#region Collection Phase

$allInventoryData = [System.Collections.Generic.List[PSCustomObject]]::new()
$perSubscriptionData = [ordered]@{}

$successCount = 0
$failCount = 0

# Connect per tenant, iterate subscriptions within each tenant
foreach ($tenant in $config.Tenants) {
    $tenantId = $tenant.TenantID
    $credPath = Join-Path $credDir $tenant.CredentialFile
    if (-not (Test-Path -Path $credPath)) {
        Write-Warning "Credential file not found: $credPath. Skipping tenant $tenantId."
        $failCount += $tenant.Subscriptions.Count
        continue
    }

    $azureConnected = $false
    try {
        Write-Host "`nConnecting to $($config.Environment) (Tenant: $tenantId)..." -ForegroundColor Gray
        $credential = Import-Clixml -Path $credPath -ErrorAction Stop
        Connect-AzAccount -Environment $config.Environment -ServicePrincipal `
            -TenantId $tenantId -Credential $credential -ErrorAction Stop | Out-Null
        $azureConnected = $true
        Write-Host "Connected to tenant $tenantId." -ForegroundColor Green
    }
    catch {
        Write-Warning "Failed to connect to tenant '$tenantId': $_"
        $failCount += $tenant.Subscriptions.Count
        continue
    }

    try {
        foreach ($sub in $tenant.Subscriptions) {
            $subName = $sub.Name
            $subAlias = if ($sub.Alias) { $sub.Alias } else { $subName.Split('-')[1] }
            Write-Host "`nProcessing Subscription: $subAlias ($subName)" -ForegroundColor Cyan

            try {
                Set-AzContext -Subscription $subName -ErrorAction Stop | Out-Null

                $vms = Get-AzVM -ErrorAction Stop | Sort-Object Name
                Write-Verbose "  Found $($vms.Count) VM(s) in $subName"

                # Cache VM sizes for this subscription's locations to avoid repeated calls
                $vmSizeCache = @{}

            $inventoryData = foreach ($vm in $vms) {
                $date = Get-Date -Format 'HH:mm'
                Write-Verbose "    Processing VM: $($vm.Name) ($date)"

                # --- Initialize all per-VM variables to prevent stale data ---
                $vmStatus = 'unknown'
                $generation = ''
                $agent = ''
                $agentStatus = 'Unknown'
                $osName = ''
                $osVersion = ''
                $bootDiag = 'none'
                $bootDiagSA = 'not configured'
                $vmRole = 'none'
                $vmDisplayName = 'none'

                # --- VM Status (includes power state, OS info, agent info) ---
                $vmStatuses = Get-AzVM -ResourceGroupName $vm.ResourceGroupName `
                    -Name $vm.Name -Status -ErrorAction Stop

                # Power state
                foreach ($status in $vmStatuses.Statuses) {
                    if ($status.Code -notlike '*provisioning*') {
                        $vmStatus = $status.Code.Split('/')[1]
                    }
                }

                # OS info from status
                $osName = $vmStatuses.OsName
                $osVersion = $vmStatuses.OsVersion
                $generation = $vmStatuses.HyperVGeneration

                # VM Agent
                $agent = $vmStatuses.VMAgent.VMAgentVersion
                if ([string]::IsNullOrEmpty($agent)) {
                    $agentStatus = 'Unknown'
                }
                else {
                    $agentStatus = $vmStatuses.VMAgent.Statuses[0].DisplayStatus
                }

                # --- Boot diagnostics ---
                $bootDiagSettings = $vm.DiagnosticsProfile.BootDiagnostics
                if ($bootDiagSettings -and $bootDiagSettings.Enabled) {
                    $storageUri = $bootDiagSettings.StorageUri
                    if ($storageUri) {
                        try {
                            $uri = [System.Uri]::new($storageUri)
                            $bootDiagSA = $uri.Host.Split('.')[0]
                            $bootDiag = 'enabled'
                        }
                        catch {
                            $bootDiag = 'enabled'
                            $bootDiagSA = 'managed'
                        }
                    }
                    else {
                        # Managed boot diagnostics (no custom storage account)
                        $bootDiag = 'enabled'
                        $bootDiagSA = 'managed'
                    }
                }

                # --- VM Size details (cached per location) ---
                $vmSizeName = $vm.HardwareProfile.VMSize
                $vmLocation = $vm.Location
                if (-not $vmSizeCache.ContainsKey($vmLocation)) {
                    $vmSizeCache[$vmLocation] = Get-AzVMSize -Location $vmLocation -ErrorAction SilentlyContinue
                }
                $vmSizeInfo = $vmSizeCache[$vmLocation] | Where-Object { $_.Name -eq $vmSizeName }

                # --- Managed disk ---
                $managed = if ($vm.StorageProfile.OsDisk.ManagedDisk) { 'Yes' } else { 'No' }

                # --- Tags (null-safe) ---
                $tags = $vm.Tags
                $appName    = if ($tags -and $tags.ContainsKey('App-name'))     { $tags['App-name'] }     else { 'None' }
                $function   = if ($tags -and $tags.ContainsKey('Function'))     { $tags['Function'] }     else { 'None' }
                $impact     = if ($tags -and $tags.ContainsKey('Impact-level')) { $tags['Impact-level'] } else { 'None' }
                $mission    = if ($tags -and $tags.ContainsKey('Mission'))      { $tags['Mission'] }      else { 'None' }
                $shutdown   = if ($tags -and $tags.ContainsKey('Shutdown'))     { $tags['Shutdown'] }     else { '' }
                $startup    = if ($tags -and $tags.ContainsKey('Startup'))      { $tags['Startup'] }      else { '' }

                # --- RBAC Role Assignments (scoped to this VM for performance) ---
                $roleAssignments = Get-AzRoleAssignment -Scope $vm.Id -ErrorAction SilentlyContinue
                if ($roleAssignments -and $roleAssignments.Count -gt 0) {
                    $vmDisplayName = ($roleAssignments.DisplayName | Where-Object { $_ }) -join '; '
                    $vmRole = ($roleAssignments.RoleDefinitionName | Where-Object { $_ }) -join '; '
                    if ([string]::IsNullOrEmpty($vmDisplayName)) { $vmDisplayName = 'none' }
                    if ([string]::IsNullOrEmpty($vmRole)) { $vmRole = 'none' }
                }

                # --- Security profile ---
                $securityType = $vm.SecurityProfile.SecurityType
                $secureBoot = $vm.SecurityProfile.UefiSettings.SecureBootEnabled
                $vtpmStatus = $vm.SecurityProfile.UefiSettings.VTpmEnabled

                # --- NICs (fetched by ResourceId for performance) ---
                $nicIds = $vm.NetworkProfile.NetworkInterfaces.Id
                $nicInfoList = @(foreach ($nicId in $nicIds) {
                    Get-AzNetworkInterface -ResourceId $nicId -ErrorAction SilentlyContinue
                })
                # Sort: primary NIC first
                $nicInfoList = @($nicInfoList | Sort-Object { -not $_.Primary })

                # Initialize NIC data for up to 3 NICs
                $nicData = @(
                    @{ Name = ''; Tag = ''; AccelNet = ''; Subnet = ''; IP1 = ''; IP2 = '' },
                    @{ Name = ''; Tag = ''; AccelNet = ''; Subnet = ''; IP1 = ''; IP2 = '' },
                    @{ Name = ''; Tag = ''; AccelNet = ''; Subnet = ''; IP1 = ''; IP2 = '' }
                )

                for ($n = 0; $n -lt [math]::Min($nicInfoList.Count, 3); $n++) {
                    $nic = $nicInfoList[$n]
                    $nicData[$n].Name = $nic.Name
                    $nicData[$n].Tag = if ($nic.Tag -and $nic.Tag.ContainsKey('App-name')) {
                        $nic.Tag['App-name']
                    } else { '' }
                    $nicData[$n].AccelNet = $nic.EnableAcceleratedNetworking

                    $primaryIpConfig = $nic.IpConfigurations | Where-Object { $_.Primary -eq $true }
                    $nicData[$n].IP1 = $primaryIpConfig.PrivateIpAddress
                    $nicData[$n].Subnet = ($primaryIpConfig.Subnet.Id | Split-Path -Leaf)

                    $secondaryIps = ($nic.IpConfigurations |
                        Where-Object { $_.Primary -ne $true }).PrivateIpAddress
                    $nicData[$n].IP2 = if ($secondaryIps) {
                        ($secondaryIps -join '; ')
                    } else { '' }
                }

                # --- Build the ordered property hashtable (49 columns) ---
                [PSCustomObject]([ordered]@{
                    'VMName'                      = $vm.Name
                    'ResourceGroupName'           = $vm.ResourceGroupName
                    'VMStatus'                    = $vmStatus
                    'Location'                    = $vm.Location
                    'App-name'                    = $appName
                    'Function'                    = $function
                    'Impact'                      = $impact
                    'Mission'                     = $mission
                    'LicenseType'                 = $vm.LicenseType
                    'VMSize'                      = $vmSizeName
                    'ManagedDisk'                 = $managed
                    'VM_Memory'                   = if ($vmSizeInfo) { $vmSizeInfo.MemoryInMB } else { '' }
                    'VM_Cores'                    = if ($vmSizeInfo) { $vmSizeInfo.NumberOfCores } else { '' }
                    'VMAgentVersion'              = $agent
                    'VM_Agent_Status'             = $agentStatus
                    'OSType'                      = $vm.StorageProfile.OsDisk.OsType
                    'OSDisk'                      = $vm.StorageProfile.OsDisk.Name
                    'OSDiskSize'                  = $vm.StorageProfile.OsDisk.DiskSizeGB
                    'BootDiag'                    = $bootDiag
                    'BootDiagSA'                  = $bootDiagSA
                    'Shutdown'                    = $shutdown
                    'Startup'                     = $startup
                    'Role'                        = $vmRole
                    'Role_Assignment'             = $vmDisplayName
                    'VM_Generation'               = $generation
                    'Nic_1_Name'                  = $nicData[0].Name
                    'Nic_1_Tag'                   = $nicData[0].Tag
                    'Nic_1_AcceleratedNetworking' = $nicData[0].AccelNet
                    'Nic_1_Subnet'               = $nicData[0].Subnet
                    'Nic_1_IP1'                   = $nicData[0].IP1
                    'Nic_1_IP2'                   = $nicData[0].IP2
                    'Nic_2_Name'                  = $nicData[1].Name
                    'Nic_2_Tag'                   = $nicData[1].Tag
                    'Nic_2_AcceleratedNetworking' = $nicData[1].AccelNet
                    'Nic_2_Subnet'               = $nicData[1].Subnet
                    'Nic_2_IP1'                   = $nicData[1].IP1
                    'Nic_2_IP2'                   = $nicData[1].IP2
                    'Nic_3_Name'                  = $nicData[2].Name
                    'Nic_3_Tag'                   = $nicData[2].Tag
                    'Nic_3_AcceleratedNetworking' = $nicData[2].AccelNet
                    'Nic_3_Subnet'               = $nicData[2].Subnet
                    'Nic_3_IP1'                   = $nicData[2].IP1
                    'Nic_3_IP2'                   = $nicData[2].IP2
                    'OS_name'                     = $osName
                    'OS_Version'                  = $osVersion
                    'SubscriptionName'            = $subAlias
                    'SecurityType'                = $securityType
                    'SecureBoot'                  = $secureBoot
                    'vTPMStatus'                  = $vtpmStatus
                })
            }

            # Store collected data
            $subInventory = @($inventoryData)
            foreach ($item in $subInventory) {
                $allInventoryData.Add($item)
            }
            $perSubscriptionData[$subAlias] = $subInventory

            Write-Host "  Collected $($subInventory.Count) VM(s)." -ForegroundColor Green
            $successCount++
        }
            catch {
                Write-Warning "Failed to process subscription '$subName': $_"
                $failCount++
            }
        }
    }
    finally {
        if ($azureConnected) {
            Write-Verbose "Disconnecting from tenant $tenantId"
            Disconnect-AzAccount -ErrorAction SilentlyContinue | Out-Null
        }
    }
}

#endregion Collection Phase

#region Build Workbook

if ($allInventoryData.Count -eq 0) {
    Write-Warning "No VM data collected from any subscription. Skipping workbook creation."
    Stop-Transcript
    return
}

Write-Host "`nBuilding workbook with $($allInventoryData.Count) total VM(s)..." -ForegroundColor Cyan

# Build filtered views
$notRunningData = @($allInventoryData | Where-Object { $_.VMStatus -ne 'running' })
$noAppNameData = @($allInventoryData | Where-Object { $_.'App-name' -eq 'None' })
$noBootDiagData = @($allInventoryData | Where-Object { $_.BootDiag -eq 'none' })
$il4Data = @($allInventoryData | Where-Object { $_.Impact -eq 'IL4' })
$il5Data = @($allInventoryData | Where-Object { $_.Impact -eq 'IL5' })

# Conditional formatting rules
$azureCfRules = @(
    # Gray: Not running VMs
    New-ConditionalText -Text 'deallocated' -Range 'C:C' -BackgroundColor LightGray
    New-ConditionalText -Text 'stopped' -Range 'C:C' -BackgroundColor LightGray
    # Red: No App-name tag
    New-ConditionalText -Text 'None' -Range 'E:E' -BackgroundColor Red -ConditionalTextColor White
    # Orange/Red: Impact level indicators
    New-ConditionalText -Text 'IL4' -Range 'G:G' -BackgroundColor Orange -ConditionalTextColor White
    New-ConditionalText -Text 'IL5' -Range 'G:G' -BackgroundColor Red -ConditionalTextColor White
    # Yellow: No managed disk
    New-ConditionalText -Text 'No' -Range 'K:K' -BackgroundColor Yellow
    # Red: VM Agent not ready
    New-ConditionalText -Text 'Unknown' -Range 'O:O' -BackgroundColor Red -ConditionalTextColor White
    # Yellow: No boot diagnostics
    New-ConditionalText -Text 'none' -Range 'S:S' -BackgroundColor Yellow
)

# Export All_VMs tab first (this creates the workbook)
$tempXlsx = Join-Path $OutputDir 'AzureInventory_All.tmp.xlsm'
if (Test-Path $tempXlsx) { Remove-Item $tempXlsx -Force }
if (Test-Path $reportFile) { Remove-Item $reportFile -Force }

$allInventoryData | Export-Excel -Path $tempXlsx -WorksheetName 'All_VMs' `
    -AutoSize -FreezePane 2, 2 -BoldTopRow -ConditionalText $azureCfRules `
    -TableName 'All_VMs' -TableStyle Medium9

# Not_Running
if ($notRunningData.Count -gt 0) {
    $notRunningData | Export-Excel -Path $tempXlsx -WorksheetName 'Not_Running' `
        -AutoSize -FreezePane 2, 2 -BoldTopRow -ConditionalText $azureCfRules `
        -TableName 'Not_Running' -TableStyle Medium9
}
else {
    Export-Excel -Path $tempXlsx -WorksheetName 'Not_Running' -InputObject $null
}

# NO_App-name
if ($noAppNameData.Count -gt 0) {
    $noAppNameData | Export-Excel -Path $tempXlsx -WorksheetName 'NO_App-name' `
        -AutoSize -FreezePane 2, 2 -BoldTopRow -ConditionalText $azureCfRules `
        -TableName 'NO_App_name' -TableStyle Medium9
}
else {
    Export-Excel -Path $tempXlsx -WorksheetName 'NO_App-name' -InputObject $null
}

# NO_BootDiag
if ($noBootDiagData.Count -gt 0) {
    $noBootDiagData | Export-Excel -Path $tempXlsx -WorksheetName 'NO_BootDiag' `
        -AutoSize -FreezePane 2, 2 -BoldTopRow -ConditionalText $azureCfRules `
        -TableName 'NO_BootDiag' -TableStyle Medium9
}
else {
    Export-Excel -Path $tempXlsx -WorksheetName 'NO_BootDiag' -InputObject $null
}

# IL4_VMs
if ($il4Data.Count -gt 0) {
    $il4Data | Export-Excel -Path $tempXlsx -WorksheetName 'IL4_VMs' `
        -AutoSize -FreezePane 2, 2 -BoldTopRow -ConditionalText $azureCfRules `
        -TableName 'IL4_VMs' -TableStyle Medium9
}
else {
    Export-Excel -Path $tempXlsx -WorksheetName 'IL4_VMs' -InputObject $null
}

# IL5_VMs
if ($il5Data.Count -gt 0) {
    $il5Data | Export-Excel -Path $tempXlsx -WorksheetName 'IL5_VMs' `
        -AutoSize -FreezePane 2, 2 -BoldTopRow -ConditionalText $azureCfRules `
        -TableName 'IL5_VMs' -TableStyle Medium9
}
else {
    Export-Excel -Path $tempXlsx -WorksheetName 'IL5_VMs' -InputObject $null
}

# Per-subscription tabs
foreach ($subAlias in $perSubscriptionData.Keys) {
    $subData = $perSubscriptionData[$subAlias]
    # Sanitize subscription name for worksheet name (max 31 chars, no special chars)
    $tabName = $subAlias -replace '[:\\/\?\*\[\]]', '_'
    if ($tabName.Length -gt 31) { $tabName = $tabName.Substring(0, 31) }

    if ($subData.Count -gt 0) {
        $subData | Export-Excel -Path $tempXlsx -WorksheetName $tabName `
            -AutoSize -FreezePane 2, 2 -BoldTopRow -ConditionalText $azureCfRules `
            -TableName ($tabName -replace '[^A-Za-z0-9_]', '_') -TableStyle Medium9
    }
    else {
        Export-Excel -Path $tempXlsx -WorksheetName $tabName -InputObject $null
    }
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
$searchWs.Cells[1, 1].Value = 'Azure VM Inventory - Search'
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
Write-Host "    Not_Running      : $($notRunningData.Count) VM(s)" -ForegroundColor Gray
Write-Host "    NO_App-name      : $($noAppNameData.Count) VM(s)" -ForegroundColor Gray
Write-Host "    NO_BootDiag      : $($noBootDiagData.Count) VM(s)" -ForegroundColor Gray
Write-Host "    IL4_VMs          : $($il4Data.Count) VM(s)" -ForegroundColor Gray
Write-Host "    IL5_VMs          : $($il5Data.Count) VM(s)" -ForegroundColor Gray
foreach ($subAlias in $perSubscriptionData.Keys) {
    Write-Host "    $subAlias : $($perSubscriptionData[$subAlias].Count) VM(s)" -ForegroundColor Gray
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
Write-Host "End of Get-AllAzureInventory.ps1 at $($endTime.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor Cyan

Stop-Transcript

#endregion Summary
