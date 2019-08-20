function Get-NetworkAdapter {
    <#
.Synopsis
Returns information on the currently used network adapters for one or more computers.
.Description
Uses Get-CimInstance to retrieve network adapter information about a computer.
.Parameter Computername
The names of one or more computers, can be piped into.
.Example
PS C:\> Get-NetworkAdapter -computername localhost
Get information about the local pc.
.Example
PS C:\> Get-ADComputer -filter * | Get-NetworkAdapter
Get information about all AD computers.
#>
    param (
        # Parameter help description
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $true)]
        $Computername
    )
    foreach ($computer in $Computername) {
        $connection = Test-Connection $Computer -Quiet -count 1
        if ($connection -eq $false) {
            Write-Information "Unable to connect to $computer." -InformationAction Continue
            continue
        }
        $output = New-Object psobject
        $NetworkAdapterConfiguration = get-ciminstance Win32_NetworkAdapterConfiguration -ComputerName $computer -filter "ipenabled = 'true'"
        foreach ($i in $NetworkAdapterConfiguration) {
            $networkadapter = Get-CimInstance CIM_NetworkAdapter -Filter "DeviceID = '$($i.index)'"
            $output | Add-Member NoteProperty "$($Computer)_Adapter$($networkadapter.DeviceID)_Name" $networkadapter.Name
            $output | Add-Member NoteProperty "$($Computer)_Adapter$($networkadapter.DeviceID)_Manufacturer" $networkadapter.Manufacturer
            $output | Add-Member NoteProperty "$($Computer)_Adapter$($networkadapter.DeviceID)_IPV4_Address" $i.Ipaddress[0]
            $output | Add-Member NoteProperty "$($Computer)_Adapter$($networkadapter.DeviceID)_IPV6_Address" $i.Ipaddress[1..(($i.ipaddress).count - 1)]
            $output | Add-Member NoteProperty "$($Computer)_Adapter$($networkadapter.DeviceID)_AdapterType" $networkadapter.AdapterType
            $output | Add-Member NoteProperty "$($Computer)_Adapter$($networkadapter.DeviceID)_MACAddress" $networkadapter.MACAddress
            $output | Add-Member NoteProperty "$($Computer)_Adapter$($networkadapter.DeviceID)_Speed" ($networkadapter.Speed / 1Gb)
        }
        trap {
            "An error occured when querying $computer`: $_"
            continue
        }
        $output
    }
}