function Get-PCInfo {
<#
.Synopsis
Returns information on one or more computer systems.
.Description
Uses Get-CimInstance to retrieve information about a computer.
.Parameter Computername
The names of one or more computers, can be piped into.
.Example
PS C:\> Get-PCInfo -computername localhost
Get information about the local pc.
.Example
PS C:\> Get-ADComputer -filter * | Get-PCInfo | Export-csv 
Get information about all AD computers and export it to csv.
#>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $True,
            ValueFromPipeline = $True)]
        $ComputerName,
        # Credentials for the CIMsession
        [Parameter(Mandatory = $false)]
        [System.Management.Automation.PSCredential]
        $Credential
    )
    foreach ($Computer in $ComputerName) {
        #Check if the pc is reachable, continue to next item if not.
        $connection = Test-Connection $Computer -Quiet -count 1
        if ($connection -eq $false) {
            Write-Information "Unable to connect to $computer." -InformationAction Continue
            continue
        }
        Write-Verbose "Collecting Information from $Computer"
        $Output = New-Object -TypeName PSObject
        #Gets the CIM info for the pc.
        $ComputerVideoCard = Get-CimInstance CIM_VideoController -ComputerName $Computer
        $PCInfo = Get-CimInstance CIM_ComputerSystem -ComputerName $Computer
        $OSInfo = Get-CimInstance CIM_OperatingSystem -ComputerName $Computer
        $ComputerProcessor = Get-CimInstance CIM_Processor -ComputerName $Computer
        #Sort wanted info into the $output variable.
        $Output | Add-Member NoteProperty "PC_Name" $PCInfo.Name
        $Output | Add-Member NoteProperty "Domain" $PCInfo.Domain
        $Output | Add-Member NoteProperty "Operating_System_Name" $OSInfo.Caption
        $Output | Add-Member NoteProperty "Operating_System_Version" $OSInfo.Version
        foreach ($Processor in $ComputerProcessor) {
            $Output | Add-Member NoteProperty "$($ComputerProcessor.DeviceID)_Name" $ComputerProcessor.Name
            $Output | Add-Member NoteProperty "$($ComputerProcessor.DeviceID)_Caption" $ComputerProcessor.Caption
            $Output | Add-Member NoteProperty "$($ComputerProcessor.DeviceID)_Vendor" $ComputerProcessor.Manufacturer
            $Output | Add-Member NoteProperty "$($ComputerProcessor.DeviceID)_Socket" $ComputerProcessor.SocketDesignation
        }
        Foreach ($Card in $ComputerVideoCard) {
            $Output | Add-Member NoteProperty "$($Card.DeviceID)_Name" $Card.Name
            $Output | Add-Member NoteProperty "$($Card.DeviceID)_Vendor" $Card.AdapterCompatibility
            $Output | Add-Member NoteProperty "$($Card.DeviceID)_PNPDeviceID" $Card.PNPDeviceID
            $Output | Add-Member NoteProperty "$($Card.DeviceID)_DriverVersion" $Card.DriverVersion
            $Output | Add-Member NoteProperty "$($Card.DeviceID)_VideoMode" $Card.VideoModeDescription
        }
        #If any errors occur continue onto next object.
        trap {
            "An error occured when querying $computer`: $_"
            continue
        }
        $Output
    }
}