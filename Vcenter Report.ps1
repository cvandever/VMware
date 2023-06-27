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
    Add-Type -AssemblyName System.Windows.Forms

    # Create an instance of the SaveFileDialog
    $fileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $fileDialog.Filter = 'CSV Files (*.csv)|*.csv'
    $fileDialog.Title = 'Save CSV File'

    # Show the file dialog and wait for user input
    $dialogResult = $fileDialog.ShowDialog()

    # Check if the user clicked the OK button
    if ($dialogResult -eq 'OK') {
        return $fileDialog.FileName
    }
    return $null
}

$ServerName = Get-ServerName
$Path = Get-SaveFileDialog

$cred = Get-Credential | Connect-VIServer -Server $ServerName  -Username $cred.Username -Password $cred.Password

$allVMs = Get-VM
$columns = 'Name', 'PowerState', 'GuestOS', 'HostName', 'ResourcePool', 'Datastore', 'vCPUs', 'MemorySize', 'ProvisionedStorage', 'PercentUsed', 'IPAddress', 'PortGroup', 'NetworkAdapter', 'AdapterType', 'MacAddress', 'SnapshotName', 'SnapshotCreated', 'VMToolsVersion'

foreach ($vm in $allVMs) {
    $networkAdapter = $vm | Get-NetworkAdapter
    $portgroup = $networkAdapter.NetworkName
    $portgroup = $portgroup -join ', ' # If there are multiple portgroups, join them with a comma
    $macAddress = $networkAdapter.MacAddress -join ', ' # If there are multiple MAC addresses, join them with a comma
    $adapterType = $networkAdapter.Type -join ', ' # If there are multiple adapter types, join them with a comma    
    $datastore = ($vm | Get-Datastore).Name -join ', ' # If there are multiple datastores, join them with a comma
    $snapshot = $vm | Get-Snapshot -ErrorAction SilentlyContinue
    $snapshotName = $null
    $snapshotCreated = $null
    if ($snapshot) {
        $snapshotName = $snapshot.Name
        $snapshotCreated = $snapshot.Created
    }

    $vmInfo = @{
        'Name' = $vm.Name
        'PowerState' = $vm.PowerState
        'GuestOS' = $vm.Guest.OSFullName
        'HostName' = $vm.VMHost.Name
        'ResourcePool' = $vm.ResourcePool.Name
        'Datastore' = $datastore
        'vCPUs' = $vm.NumCpu
        'MemorySize' = $vm.MemoryGB
        'ProvisionedStorage' = [Math]::Round($vm.ProvisionedSpaceGB,0)
        'PercentUsed' = [Math]::Round(($vm.UsedSpaceGB / $vm.ProvisionedSpaceGB * 100),0)
        'IPAddress' = $vm.Guest.IPAddress -join ', '
        'PortGroup' = $portgroup
        'NetworkAdapter' = $networkAdapter.NetworkName -join ', '
        'AdapterType' = $adapterType
        'MacAddress' = $macAddress
        'SnapshotName' = $snapshotName
        'SnapshotCreated' = $snapshotCreated
        'VMToolsVersion' = $vm.Guest.ToolsVersion
    }
    

    $vmInfoObject = New-Object -TypeName PSObject -Property $vmInfo
    $selectedProperties = $vmInfoObject | Select-Object -Property $columns
    $selectedProperties | Export-Csv -Path $Path -Append -NoTypeInformation
}
