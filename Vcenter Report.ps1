# Install the ImportExcel module if not already installed
if (-not (Get-Module -Name ImportExcel -ListAvailable)) {
    Install-Module -Name ImportExcel -Scope CurrentUser -Force
}


function Get-ServerName {
    param (
        [string]$DefaultServerName = 'spolabvcsa.lab.ceriumnetworks.com'
    )

    $serverName = Read-Host "Enter the server name (default: $DefaultServerName)"

    if ([string]::IsNullOrWhiteSpace($serverName)) {
        $serverName = $DefaultServerName
    }

    return $serverName
}

function Get-SaveFileDialog {
    param (
        [string]$ServerName
    )

    Add-Type -AssemblyName System.Windows.Forms
    # Get the current date and time
    $currentDateTime = Get-Date -Format 'MMddyy_HHmm'
    # Get the current working directory
    $defaultFolder = (Get-Item -Path ".\").FullName
    # Create the default filename using the server name and current date/time
    $defaultFileName = "$ServerName" + "_$currentDateTime.xlsx"

    # Create an instance of the SaveFileDialog
    $fileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $fileDialog.Filter = 'Excel Files (*.xlsx)|*.xlsx'
    $fileDialog.Title = 'Save Excel File'
    $fileDialog.InitialDirectory = $DefaultFolder
    $fileDialog.FileName = $defaultFileName

    # Show the file dialog and wait for user input
    $dialogResult = $fileDialog.ShowDialog()

    # Check if the user clicked the OK button
    if ($dialogResult -eq 'OK') {
        return $fileDialog.FileName
    }
    return $null
}

function Get-VMNetworkInfo {
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        $VM
    )

    process {
        $networkAdapter = $VM | Get-NetworkAdapter
        $connected = $networkAdapter.ConnectionState.Connected -join ', '
        $portgroup = $networkAdapter.NetworkName -join ', '
        $macAddress = $networkAdapter.MacAddress -join ', '
        $adapterType = $networkAdapter.Type -join ', '

        [PSCustomObject]@{
            PortGroup = $portgroup
            ConnectionState = $connected
            AdapterType = $adapterType
            MacAddress = $macAddress
        }
    }
}

function Get-VMSnapshotInfo {
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        $VM
    )

    process {
        $snapshots = $VM | Get-Snapshot

        if ($snapshots) {
            $snapshotNames = $snapshots.Name -join ', '
            $snapshotCreated = $snapshots.Created.Date -join ', '
            if ($snapshots.SizeGB -gt 0.01) {
                $CumSnapshotSize = "{0:N2}GB" -f ($snapshots | Measure-Object -Property SizeGB -Sum).Sum
            }
            else {
                $CumSnapshotSize = "~0.01GB"
            }

            [PSCustomObject]@{
                SnapshotName = $snapshotNames
                SnapshotCreated = $snapshotCreated
                CumulativeSnapshotSize = $CumSnapshotSize
            }
        }
    }
}

function Get-VMEventUsers {
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        $VM
    )

    process {
        $events = Get-VIEvent -Entity $VM -MaxSamples ([int]::MaxValue)
        $users = $events | Where-Object { $_.UserName -like "CERIUMNETWORKS\*" } | ForEach-Object {
            $_.UserName -replace "^CERIUMNETWORKS\\"
        } | Select-Object -Unique
        return $users -join ', '
    }
}

# Retrieve Vsphere credentials and servername from user input
# Open File Dialog to select location to save report
$creds = Get-Credential 
$ServerName = Get-ServerName
$Path = Get-SaveFileDialog -ServerName $ServerName

Connect-VIServer -Server $ServerName  -Credential $creds

$allVMs = Get-VM

# Create an array to store the worksheets
$worksheets = @()

# If you add more columns, make sure to add them to the $columns variable below
# Creates ordered hashtable with column names and values. In order left to right
$columnsGeneralInfo = ('Name', 'PowerState', 'GuestOS', 'HostName','ResourcePool',
    'Datastore', 'vCPUs', 'MemorySize', 'ProvisionedStorage', 'PercentStorageUsed',
    'VMToolsVersion', 'VMOwner', 'VMOwnerTeam', 'VMCreation', 'VMExpiration', 'RecentUsers' )


$columnsNetworkInfo = ('Name', 'IPAddress', 'PortGroup', 'ConnectionState', 'AdapterType', 'MacAddress')


$columnsSnapshotInfo = ('Name', 'SnapshotName', 'SnapshotCreated', 'CumulativeSnapshotSize')


foreach ($vm in $allVMs) {
    $networkInfo = $vm | Get-VMNetworkInfo
    $snapshotInfo = $vm | Get-VMSnapshotInfo
    #$users = Get-VMEventUsers -VM $vm
    
    $generalInfoData = [PSCustomObject]@{
        Name = $vm.Name
        PowerState = $vm.PowerState
        GuestOS = $vm.Guest.OSFullName
        HostName = $vm.VMHost.Name
        ResourcePool = $vm.ResourcePool.Name
        Datastore = ($vm | Get-Datastore).Name -join ', '
        vCPUs = $vm.NumCpu
        MemorySize = $vm.MemoryGB
        ProvisionedStorage = "{0:N2}GB" -f $vm.ProvisionedSpaceGB
        PercentStorageUsed = "{0:N2}%" -f ($vm.UsedSpaceGB / $vm.ProvisionedSpaceGB * 100)
        VMToolsVersion = $vm.Guest.ToolsVersion
        VMOwner = $vm.CustomFields["VM Owner"]
        VMOwnerTeam = $vm.CustomFields["Owner's Team"]
        VMCreation = $vm.CustomFields["Creation Date"]
        VMExpiration = $vm.CustomFields["Expiration Date"]
        #RecentUsers = $users
    }
    $networkInfoData = [PSCustomObject]@{
        Name = $vm.Name
        IPAddress = $vm.Guest.IPAddress -join ', '
        PortGroup = $networkInfo.PortGroup
        ConnectionState = $networkInfo.ConnectionState
        AdapterType = $networkInfo.AdapterType
        MacAddress = $networkInfo.MacAddress
    }
        
    $snapshotInfoData = [PSCustomObject]@{
        Name = $vm.Name
        SnapshotName = $snapshotInfo.SnapshotName
        SnapshotCreated = $snapshotInfo.SnapshotCreated
        CumulativeSnapshotSize = $snapshotInfo.CumulativeSnapshotSize    
    }

    $worksheets += @{
        Name = 'General System Info'
        Data = $generalInfoData
        Columns = $columnsGeneralInfo
    }

    $worksheets += @{
        Name = 'Network Info'
        Data = $networkInfoData
        Columns = $columnsNetworkInfo
    }

    $worksheets += @{
        Name = 'Snapshot Info'
        Data = $snapshotInfoData
        Columns = $columnsSnapshotInfo
    }
}
    
Disconnect-VIServer -Server $ServerName -Confirm:$false
Write-Host "Disconnected from $ServerName"
Write-Host "Exporting data to $Path"

ForEach ($worksheet in $worksheets) {
    $worksheetName = $worksheet.Name
    $worksheetData = $worksheet.Data
    $worksheetColumns = $worksheet.Columns

    $worksheetData | Select-Object -Property $worksheetColumns | Export-Excel -Path $Path -WorksheetName $worksheetName -AutoSize -TableStyle 'Medium1' `
        -BoldTopRow -FreezeTopRow -AutoFilter -AutoNameRange -TableName "Table $worksheetName" `
        -Append 
}

