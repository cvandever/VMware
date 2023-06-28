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
        [string]$ServerName,
        [string]$DefaultFolder = 'C:\Reports'
    )

    Add-Type -AssemblyName System.Windows.Forms

    # Get the current date and time
    $currentDateTime = Get-Date -Format 'MMddyy_HHmm'

    # Get the current working directory
    $defaultFolder = (Get-Item -Path ".\").FullName

    # Create the default filename using the server name and current date/time
    $defaultFileName = "$ServerName" + "_$currentDateTime.csv"

    # Create an instance of the SaveFileDialog
    $fileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $fileDialog.Filter = 'CSV Files (*.csv)|*.csv'
    $fileDialog.Title = 'Save CSV File'
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
        $portgroup = $networkAdapter.NetworkName -join ', '
        $macAddress = $networkAdapter.MacAddress -join ', '
        $adapterType = $networkAdapter.Type -join ', '

        [PSCustomObject]@{
            PortGroup = $portgroup
            NetworkAdapter = $networkAdapter.NetworkName
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
            $snapshotSize = "{0:N2}" -f ($snapshots.SizeGB -join ', ')

            [PSCustomObject]@{
                SnapshotName = $snapshotNames
                SnapshotCreated = $snapshotCreated
                SnapshotSize = $snapshotSize
            }
        } else {
            [PSCustomObject]@{
                SnapshotName = $null
                SnapshotCreated = $null
                SnapshotSize = $null
            }
        }
    }
}

# Retrieve Vsphere credentials and servername from user input
# Open File Dialog to select location to save report
$creds = Get-Credential 
$ServerName = Get-ServerName
$Path = Get-SaveFileDialog -ServerName $ServerName

Connect-VIServer -Server $ServerName  -Credential $creds

$allVMs = Get-VM

# If you add more columns, make sure to add them to the $columns variable below
# Creates ordered hashtable with column names and values. In order left to right
$columns = ('Name', 'PowerState', 'GuestOS', 'HostName','ResourcePool',
    'Datastore', 'vCPUs', 'MemorySize', 'ProvisionedStorage', 'PercentUsed',
    'IPAddress', 'PortGroup', 'NetworkAdapter', 'AdapterType', 'MacAddress',
    'SnapshotName', 'SnapshotCreated', 'SnapshotSize', 'VMToolsVersion',
    'VMOwner', 'VMOwnerTeam', 'VMExpiration')

    $exportData = foreach ($vm in $allVMs) {
        $networkInfo = $vm | Get-VMNetworkInfo
        $snapshotInfo = $vm | Get-VMSnapshotInfo
    
    [PSCustomObject]@{
        Name = $vm.Name
        PowerState = $vm.PowerState
        GuestOS = $vm.Guest.OSFullName
        HostName = $vm.VMHost.Name
        ResourcePool = $vm.ResourcePool.Name
        Datastore = ($vm | Get-Datastore).Name -join ', '
        vCPUs = $vm.NumCpu
        MemorySize = $vm.MemoryGB
        ProvisionedStorage = "{0:N2}" -f $vm.ProvisionedSpaceGB
        PercentUsed = "{0:N2}" -f ($vm.UsedSpaceGB / $vm.ProvisionedSpaceGB * 100)
        IPAddress = $vm.Guest.IPAddress -join ', '
        PortGroup = $networkInfo.PortGroup
        NetworkAdapter = $networkInfo.NetworkAdapter
        AdapterType = $networkInfo.AdapterType
        MacAddress = $networkInfo.MacAddress
        SnapshotName = $snapshotInfo.SnapshotName
        SnapshotCreated = $snapshotInfo.SnapshotCreated
        SnapshotSize = $snapshotInfo.SnapshotSize
        VMToolsVersion = $vm.Guest.ToolsVersion
        VMOwner = $vm.CustomFields["VM Owner"] -join ', '
        VMOwnerTeam = $vm.CustomFields["VM Owner Team"] -join ', '
        VMExpiration = $vm.CustomFields["VM Expiration"] -join ', '
    }
}
    
Disconnect-VIServer -Server $ServerName -Confirm:$false
Write-Host "Disconnected from $ServerName"
Write-Host "Exporting data to $Path"

$selectedProperties = $exportData | Select-Object -Property $columns
$selectedProperties | Export-Csv -Path $Path -NoTypeInformation -Append

