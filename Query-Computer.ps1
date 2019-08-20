# Loading external assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#Sets the powershell session to administrator if user isn't admin.
$CurrentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())

if (-not($CurrentUser.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator))) {
    $argumentlist = "-noprofile -noexit -file ($myInvocation.Mycommand.Definition)"
    Start-Process powershell.exe -verb RunAs -WindowStyle Hidden -ArgumentList $argumentlist
}

#Loads the main form.
$FMMain = New-Object System.Windows.Forms.Form
$wshell = New-Object -ComObject Wscript.Shell

#Loads the form objects.
$TBName = New-Object System.Windows.Forms.TextBox
$Label = New-Object System.Windows.Forms.Label
$BTGetInfo = New-Object System.Windows.Forms.Button
$BTClear = New-Object System.Windows.Forms.Button
$tabControl1 = New-Object System.Windows.Forms.TabControl
$OSPage = New-Object System.Windows.Forms.TabPage
$pcPage = New-Object System.Windows.Forms.TabPage
$CPUPage = New-Object System.Windows.Forms.TabPage
$GPUPage = New-Object System.Windows.Forms.TabPage
$NetPage = New-Object System.Windows.Forms.TabPage
$DGVPCInfo = New-Object System.Windows.Forms.DataGridView
$DGVOSInfo = New-Object System.Windows.Forms.DataGridView
$DGVCPUInfo = New-Object System.Windows.Forms.DataGridView
$DGVGPUInfo = New-Object System.Windows.Forms.DataGridView
$DGVNetInfo = New-Object System.Windows.Forms.DataGridView
$ContextMenuStripPC = New-Object System.Windows.Forms.ContextMenuStrip
$ContextMenuStripOS = New-Object System.Windows.Forms.ContextMenuStrip
$ContextMenuStripCPU = New-Object System.Windows.Forms.ContextMenuStrip
$ContextMenuStripGPU = New-Object System.Windows.Forms.ContextMenuStrip
$ContextMenuStripNet = New-Object System.Windows.Forms.ContextMenuStrip
#
#Text Boxes listed here
# TBName
#
$TBName.Location = New-Object System.Drawing.Point(60, 24)
$TBName.Name = "TBName"
$TBName.Size = New-Object System.Drawing.Size(217, 20)
$TBName.TabIndex = 0
#
#Context menu objects listed here
# ContextMenuStripPC
#
#Each context menu copies an indexed item from a selected row.
$ContextMenuStripPC.Items.Add("Copy Domain").add_click( { Set-Clipboard $DGVPCInfo.SelectedRows[1].Value })
$ContextMenuStripPC.Items.Add("Copy Motherboard").add_click( { Set-Clipboard $DGVPCInfo.SelectedRows[2].Value })
$ContextMenuStripPC.Items.Add("Copy Bios_Version").add_click( { Set-Clipboard $DGVPCInfo.SelectedRows[3].Value })
$ContextMenuStripPC.Items.Add("Copy SerialNumber").add_click( { Set-Clipboard $DGVPCInfo.SelectedRows[4].Value })
#
# ContextMenuStripOS
#
$ContextMenuStripOS.Items.Add("Copy OS_Type").add_click( { Set-Clipboard $DGVOSInfo.SelectedRows[1].Value })
$ContextMenuStripOS.Items.Add("Copy OS_Version").add_click( { Set-Clipboard $DGVOSInfo.SelectedRows[2].Value })
$ContextMenuStripOS.Items.Add("Copy SystemDrive").add_click( { Set-Clipboard $DGVOSInfo.SelectedRows[3].Value })
$ContextMenuStripOS.Items.Add("Copy SerialNumber").add_click( { Set-Clipboard $DGVOSInfo.SelectedRows[4].Value })
#
# ContextMenuStripCPU
#
$ContextMenuStripCPU.Items.Add("Copy Name").add_click( { Set-Clipboard $DGVCPUInfo.SelectedRows[1].Value })
$ContextMenuStripCPU.Items.Add("Copy DeviceID").add_click( { Set-Clipboard $DGVCPUInfo.SelectedRows[2].Value })
$ContextMenuStripCPU.Items.Add("Copy Caption").add_click( { Set-Clipboard $DGVCPUInfo.SelectedRows[3].Value })
$ContextMenuStripCPU.Items.Add("Copy Vendor").add_click( { Set-Clipboard $DGVCPUInfo.SelectedRows[4].Value })
$ContextMenuStripCPU.Items.Add("Copy Socket").add_click( { Set-Clipboard $DGVCPUInfo.SelectedRows[5].Value })
#
# ContextMenuStripGPU
#
$ContextMenuStripGPU.Items.Add("Copy Card_Name").add_click( { Set-Clipboard $DGVGPUInfo.SelectedRows[1].Value })
$ContextMenuStripGPU.Items.Add("Copy Card_Vendor").add_click( { Set-Clipboard $DGVGPUInfo.SelectedRows[2].Value })
$ContextMenuStripGPU.Items.Add("Copy Card_PNPDeviceID").add_click( { Set-Clipboard $DGVGPUInfo.SelectedRows[3].Value })
$ContextMenuStripGPU.Items.Add("Copy Card_DriverVersion").add_click( { Set-Clipboard $DGVGPUInfo.SelectedRows[4].Value })
$ContextMenuStripGPU.Items.Add("Copy Card_VideoMode").add_click( { Set-Clipboard $DGVGPUInfo.SelectedRows[5].Value })
#
# ContextMenuStripNet
#
$ContextMenuStripNet.Items.Add("Copy Adapter_Name").add_click( { Set-Clipboard $DGVNetInfo.SelectedRows[1].Value })
$ContextMenuStripNet.Items.Add("Copy Adapter_IPV4").add_click( { Set-Clipboard $DGVNetInfo.SelectedRows[2].Value })
$ContextMenuStripNet.Items.Add("Copy Adapter_AdapterType").add_click( { Set-Clipboard $DGVNetInfo.SelectedRows[3].Value })
$ContextMenuStripNet.Items.Add("Copy Adapter_MACAddress").add_click( { Set-Clipboard $DGVNetInfo.SelectedRows[4].Value })
$ContextMenuStripNet.Items.Add("Copy Adapter_Speed").add_click( { Set-Clipboard $DGVNetInfo.SelectedRows[5].Value })
#
# Label objects listed here
# Label
#
$Label.AutoSize = $true
$Label.Enabled = $false
$Label.Location = New-Object System.Drawing.Point(57, 8)
$Label.Name = "Label"
$Label.Size = New-Object System.Drawing.Size(80, 13)
$Label.TabIndex = 10
$Label.Text = "ComputerName"
#
# Button objects listed here
# BTGetInfo
#
$BTGetInfo.Location = New-Object System.Drawing.Point(331, 24)
$BTGetInfo.Name = "BTGetInfo"
$BTGetInfo.Size = New-Object System.Drawing.Size(75, 20)
$BTGetInfo.TabIndex = 1
$BTGetInfo.Text = "Get Info"
$BTGetInfo.UseVisualStyleBackColor = $true
$BTGetInfo.add_click( {
        # Set computer name based on whether text box was filled in.
        $computername = if ($TBName.Text -gt 0) {
            $TBName.Text
        }
        else {
            $env:COMPUTERNAME
        }
        # Check if computer is reachable if not show error and add to log file.
        if ((Test-Connection -ComputerName $computername -Quiet -Count 1) -eq $false) {
            $wshell.Popup("Unable to connect to $computername. `nDetails can be found at $env:USERPROFILE\Query-Computer log.txt.", 0, "Connection Error", 0x0)
            Test-Connection -ComputerName $computername -Count 1 -ErrorAction SilentlyContinue
            logAction "$($error | select-object -first 1)"
        }
        else {
            # Just run each function.
            Get-OSInfo
            Get-PCInfo
            Get-CPUInfo
            Get-GPUInfo
            Get-NetInfo
        }
        $TBName.Clear()
    })
#
# BTClear
#
$BTClear.Location = New-Object System.Drawing.Point(424, 24)
$BTClear.Name = "BTClear"
$BTClear.Size = New-Object System.Drawing.Size(75, 20)
$BTClear.TabIndex = 2
$BTClear.Text = "Clear Table"
$BTClear.UseVisualStyleBackColor = $true
$BTClear.add_click( {
        # Clear all the rows in each DataGridView
        $DGVOSInfo.Rows.Clear()
        $DGVPCInfo.Rows.Clear()
        $DGVCPUInfo.Rows.Clear()
        $DGVGPUInfo.Rows.Clear()
        $DGVNetInfo.Rows.Clear()
    })
#
# tabControl1
#
$tabControl1.Controls.Add($pcPage)
$tabControl1.Controls.Add($OSPage)
$tabControl1.Controls.Add($CPUPage)
$tabControl1.Controls.Add($GPUPage)
$tabControl1.Controls.Add($NetPage)
$tabControl1.Location = New-Object System.Drawing.Point(12, 80)
$tabControl1.Name = "tabControl1"
$tabControl1.SelectedIndex = 0
$tabControl1.Size = New-Object System.Drawing.Size(778, 415)
$tabControl1.TabIndex = 4
#
# OSPage
#
$OSPage.Controls.Add($DGVOSInfo)
$OSPage.Location = New-Object System.Drawing.Point(4, 22)
$OSPage.Name = "OSPage"
$OSPage.Padding = New-Object System.Windows.Forms.Padding(3)
$OSPage.Size = New-Object System.Drawing.Size(770, 389)
$OSPage.TabIndex = 1
$OSPage.Text = "OS Info"
$OSPage.UseVisualStyleBackColor = $true
#
# pcPage
#
$pcPage.Controls.Add($DGVPCInfo)
$pcPage.Location = New-Object System.Drawing.Point(4, 22)
$pcPage.Name = "pcPage"
$pcPage.Padding = New-Object System.Windows.Forms.Padding(3)
$pcPage.Size = New-Object System.Drawing.Size(770, 389)
$pcPage.TabIndex = 0
$pcPage.Text = "PC Info"
$pcPage.UseVisualStyleBackColor = $true
#
# CPUPage
#
$CPUPage.Controls.Add($DGVCPUInfo)
$CPUPage.Location = New-Object System.Drawing.Point(4, 22)
$CPUPage.Name = "CPUPage"
$CPUPage.Size = New-Object System.Drawing.Size(770, 389)
$CPUPage.TabIndex = 2
$CPUPage.Text = "CPU Info"
$CPUPage.UseVisualStyleBackColor = $true
#
# GPUPage
#
$GPUPage.Controls.Add($DGVGPUInfo)
$GPUPage.Location = New-Object System.Drawing.Point(4, 22)
$GPUPage.Name = "GPUPage"
$GPUPage.Size = New-Object System.Drawing.Size(770, 389)
$GPUPage.TabIndex = 3
$GPUPage.Text = "GPU Info"
$GPUPage.UseVisualStyleBackColor = $true
#
# NetPage
#
$NetPage.Controls.Add($DGVNetInfo)
$NetPage.Location = New-Object System.Drawing.Point(4, 22)
$NetPage.Name = "NetPage"
$NetPage.Size = New-Object System.Drawing.Size(770, 389)
$NetPage.TabIndex = 4
$NetPage.Text = "Net Info"
$NetPage.UseVisualStyleBackColor = $true
#
# DGVPCInfo
#
$DGVPCInfo.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::AutoSize
$DGVPCInfo.Location = New-Object System.Drawing.Point(6, 6)
$DGVPCInfo.Name = "DGVPCInfo"
$DGVPCInfo.Size = New-Object System.Drawing.Size(758, 377)
$DGVPCInfo.TabIndex = 0
$DGVPCInfo.RowHeadersVisible = $false
$DGVPCInfo.AutoSizeColumnsMode = 'Fill'
$DGVPCInfo.AllowUserToResizeRows = $false
$DGVPCInfo.selectionmode = 'FullRowSelect'
$DGVPCInfo.MultiSelect = $true
$DGVPCInfo.AllowUserToAddRows = $false
$DGVPCInfo.ReadOnly = $true
$DGVPCInfo.ColumnCount = 5
$DGVPCInfo.ColumnHeadersVisible = $true
$DGVPCInfo.Columns[0].Name = "PC_Name"
$DGVPCInfo.Columns[0].Width = 100
$DGVPCInfo.Columns[1].Name = "Domain"
$DGVPCInfo.Columns[2].Name = "Motherboard"
$DGVPCInfo.Columns[2].Width = 70
$DGVPCInfo.Columns[3].Name = "Bios_Version"
$DGVPCInfo.Columns[3].Width = 90
$DGVPCInfo.Columns[4].Name = "Motherboard_SerialNumber"
$DGVPCInfo.Columns[4].Width = 160
$DGVPCInfo.ContextMenuStrip = $ContextMenuStripPC
$DGVPCInfo.add_doubleclick( {
        Remove-Row
    })
#
# DGVOSInfo
#
$DGVOSInfo.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::AutoSize
$DGVOSInfo.Location = New-Object System.Drawing.Point(6, 6)
$DGVOSInfo.Name = "DGVOSInfo"
$DGVOSInfo.Size = New-Object System.Drawing.Size(758, 377)
$DGVOSInfo.TabIndex = 1
$DGVOSInfo.RowHeadersVisible = $false
$DGVOSInfo.AutoSizeColumnsMode = 'Fill'
$DGVOSInfo.AllowUserToResizeRows = $false
$DGVOSInfo.selectionmode = 'FullRowSelect'
$DGVOSInfo.MultiSelect = $true
$DGVOSInfo.AllowUserToAddRows = $false
$DGVOSInfo.ReadOnly = $true
$DGVOSInfo.ColumnCount = 5
$DGVOSInfo.ColumnHeadersVisible = $true
$DGVOSInfo.Columns[0].Name = "PC_Name"
$DGVOSInfo.Columns[0].Width = 100
$DGVOSInfo.Columns[1].Name = "OS_Type"
$DGVOSInfo.Columns[2].Name = "OS_Version"
$DGVOSInfo.Columns[2].Width = 70
$DGVOSInfo.Columns[3].Name = "OS_SystemDrive"
$DGVOSInfo.Columns[3].Width = 90
$DGVOSInfo.Columns[4].Name = "OS_SerialNumber"
$DGVOSInfo.Columns[4].Width = 160
$DGVOSInfo.ContextMenuStrip = $ContextMenuStripOS
$DGVOSInfo.add_doubleclick( {
        Remove-Row
    })
#
# DGVCPUInfo
#
$DGVCPUInfo.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::AutoSize
$DGVCPUInfo.Location = New-Object System.Drawing.Point(6, 6)
$DGVCPUInfo.Name = "DGVCPUInfo"
$DGVCPUInfo.Size = New-Object System.Drawing.Size(758, 377)
$DGVCPUInfo.TabIndex = 1
$DGVCPUInfo.RowHeadersVisible = $false
$DGVCPUInfo.AutoSizeColumnsMode = 'Fill'
$DGVCPUInfo.AllowUserToResizeRows = $false
$DGVCPUInfo.selectionmode = 'FullRowSelect'
$DGVCPUInfo.MultiSelect = $true
$DGVCPUInfo.AllowUserToAddRows = $false
$DGVCPUInfo.ReadOnly = $true
$DGVCPUInfo.ColumnCount = 6
$DGVCPUInfo.ColumnHeadersVisible = $true
$DGVCPUInfo.Columns[0].Name = "PC_Name"
$DGVCPUInfo.Columns[0].Width = 100
$DGVCPUInfo.Columns[1].Name = "Processor_Name"
$DGVCPUInfo.Columns[2].Name = "Processor_ID"
$DGVCPUInfo.Columns[2].Width = 70
$DGVCPUInfo.Columns[3].Name = "Processor_Caption"
$DGVCPUInfo.Columns[3].Width = 90
$DGVCPUInfo.Columns[4].Name = "Processor_Vendor"
$DGVCPUInfo.Columns[4].Width = 160
$DGVCPUInfo.Columns[5].Name = "Processor_Socket"
$DGVCPUInfo.ContextMenuStrip = $ContextMenuStripCPU
$DGVCPUInfo.add_doubleclick( {
        Remove-Row
    })
#
# DGVGPUInfo
#
$DGVGPUInfo.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::AutoSize
$DGVGPUInfo.Location = New-Object System.Drawing.Point(6, 6)
$DGVGPUInfo.Name = "DGVGPUInfo"
$DGVGPUInfo.Size = New-Object System.Drawing.Size(758, 377)
$DGVGPUInfo.TabIndex = 1
$DGVGPUInfo.RowHeadersVisible = $false
$DGVGPUInfo.AutoSizeColumnsMode = 'Fill'
$DGVGPUInfo.AllowUserToResizeRows = $false
$DGVGPUInfo.selectionmode = 'FullRowSelect'
$DGVGPUInfo.MultiSelect = $true
$DGVGPUInfo.AllowUserToAddRows = $false
$DGVGPUInfo.ReadOnly = $true
$DGVGPUInfo.ColumnCount = 6
$DGVGPUInfo.ColumnHeadersVisible = $true
$DGVGPUInfo.Columns[0].Name = "PC_Name"
$DGVGPUInfo.Columns[0].Width = 100
$DGVGPUInfo.Columns[1].Name = "Card_Name"
$DGVGPUInfo.Columns[2].Name = "Card_Vendor"
$DGVGPUInfo.Columns[2].Width = 70
$DGVGPUInfo.Columns[3].Name = "Card_PNPDeviceID"
$DGVGPUInfo.Columns[3].Width = 90
$DGVGPUInfo.Columns[4].Name = "Card_DriverVersion"
$DGVGPUInfo.Columns[4].Width = 160
$DGVGPUInfo.Columns[5].Name = "Card_VideoMode"
$DGVGPUInfo.ContextMenuStrip = $ContextMenuStripGPU
$DGVGPUInfo.add_doubleclick( {
        Remove-Row
    })
#
# DGVNetInfo
#
$DGVNetInfo.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::AutoSize
$DGVNetInfo.Location = New-Object System.Drawing.Point(6, 6)
$DGVNetInfo.Name = "DGVNetInfo"
$DGVNetInfo.Size = New-Object System.Drawing.Size(758, 377)
$DGVNetInfo.TabIndex = 2
$DGVNetInfo.RowHeadersVisible = $false
$DGVNetInfo.AutoSizeColumnsMode = 'Fill'
$DGVNetInfo.AllowUserToResizeRows = $false
$DGVNetInfo.selectionmode = 'FullRowSelect'
$DGVNetInfo.MultiSelect = $true
$DGVNetInfo.AllowUserToAddRows = $false
$DGVNetInfo.ReadOnly = $true
$DGVNetInfo.ColumnCount = 6
$DGVNetInfo.ColumnHeadersVisible = $true
$DGVNetInfo.Columns[0].Name = "PC_Name"
$DGVNetInfo.Columns[0].Width = 100
$DGVNetInfo.Columns[1].Name = "NetAdapter_Name"
$DGVNetInfo.Columns[2].Name = "NetAdapter_IPv4_Address"
$DGVNetInfo.Columns[2].Width = 70
$DGVNetInfo.Columns[3].Name = "NetAdapter_AdapterType"
$DGVNetInfo.Columns[3].Width = 90
$DGVNetInfo.Columns[4].Name = "NetAdapter_MACAddress"
$DGVNetInfo.Columns[4].Width = 160
$DGVNetInfo.Columns[5].Name = "NetAdapter_Speed(GB)"
$DGVNetInfo.ContextMenuStrip = $ContextMenuStripNet
$DGVNetInfo.add_doubleclick( {
        Remove-Row
    })
#
# FMMain
#
$FMMain.ClientSize = New-Object System.Drawing.Size(802, 507)
$FMMain.Controls.Add($tabControl1)
$FMMain.Controls.Add($BTClear)
$FMMain.Controls.Add($BTGetInfo)
$FMMain.Controls.Add($Label)
$FMMain.Controls.Add($TBName)
$FMMain.Name = "FMMain"
$FMMain.Text = "Query-Computer"

function Get-OSInfo {
    # Use CIM to get data on the OS.
    $Output = New-Object -TypeName PSObject
    $OS = Get-CIMInstance CIM_operatingsystem -computername $computername
    $Output | Add-Member NoteProperty "PC_Name" $computername
    $Output | Add-Member NoteProperty "OS_Type" $OS.Caption
    $Output | Add-Member NoteProperty "OS_Version" $OS.Version
    $Output | Add-Member NoteProperty "OS_SystemDrive" $OS.SystemDrive
    $Output | Add-Member NoteProperty "OS_SerialNumber" $OS.SerialNumber
    $DGVOSInfo.Rows.Add(($Output).PC_Name, ($Output).OS_Type, ($Output).OS_Version, ($Output).OS_SystemDrive, ($Output).OS_SerialNumber)
}

function Get-PCInfo {
    # Use CIM to get data on the PC
    $Output = New-Object -TypeName PSObject
    $PC = Get-CIMInstance CIM_computersystem -computername $computername
    $bios = Get-CimInstance Win32_BIOS -ComputerName $computername
    $Output | Add-Member NoteProperty "PC_Name" $computername
    $Output | Add-Member NoteProperty "Domain" $PC.Domain
    $Output | Add-Member NoteProperty "Motherboard" $PC.Model
    $Output | Add-Member NoteProperty "Bios_Version" $bios.Version
    $Output | Add-Member NoteProperty "Motherboard_SerialNumber" $Bios.SerialNumber
    $DGVPCInfo.Rows.Add(($Output).PC_Name, ($Output).Domain, ($Output).Motherboard, ($Output).Bios_Version, ($Output).Motherboard_SerialNumber)
}

function Get-CPUInfo {
    # Use CIM to get data on the CPU
    $Output = New-Object -TypeName PSObject
    $ComputerProcessor = Get-CimInstance CIM_Processor -ComputerName $Computer
    foreach ($Processor in $ComputerProcessor) {
        # Splits up CPU data into each CPU for better layout
        $Output | Add-Member NoteProperty "PC_Name" $computername
        $Output | Add-Member NoteProperty "Processor_Name" $ComputerProcessor.Name
        $Output | Add-Member NoteProperty "Processor_DeviceID" $ComputerProcessor.DeviceID
        $Output | Add-Member NoteProperty "Processor_Caption" $ComputerProcessor.Caption
        $Output | Add-Member NoteProperty "Processor_Vendor" $ComputerProcessor.Manufacturer
        $Output | Add-Member NoteProperty "Processor_Socket" $ComputerProcessor.SocketDesignation
        $DGVCPUInfo.Rows.Add(($Output).PC_Name, ($Output).Processor_Name, ($Output).Processor_DeviceID, ($Output).Processor_Caption, ($Output).Processor_Vendor, ($Output).Processor_Socket)
    }
}

function Get-GPUInfo {
    # Use CIM to get data on the GPU
    $Output = New-Object -TypeName PSObject
    $ComputerVideoCard = Get-CimInstance CIM_VideoController -ComputerName $computername
    foreach ($Card in $ComputerVideoCard) {
        # Splits up GPU data into each GPU for better layout
        $Output | Add-Member NoteProperty "PC_Name" $computername
        $Output | Add-Member NoteProperty "Card_Name" $Card.Name
        $Output | Add-Member NoteProperty "Card_Vendor" $Card.AdapterCompatibility
        $Output | Add-Member NoteProperty "Card_PNPDeviceID" $Card.PNPDeviceID
        $Output | Add-Member NoteProperty "Card_DriverVersion" $Card.DriverVersion
        $Output | Add-Member NoteProperty "Card_VideoMode" $Card.VideoModeDescription
        $DGVGPUInfo.Rows.Add(($Output).PC_Name, ($Output).Card_Name, ($Output).Card_Vendor, ($Output).Card_PNPDeviceID, ($Output).Card_DriverVersion, ($Output).Card_VideoMode)
    }
}

function Get-NetInfo {
    # Use CIM to get data on the Network Adapters
    $NetworkAdapterConfiguration = get-ciminstance Win32_NetworkAdapterConfiguration -ComputerName $computername -filter "ipenabled = 'true'"
    foreach ($Adapter in $NetworkAdapterConfiguration) {
        # Splits up adapter data into each adapter for better layout
        $Output = New-Object -TypeName PSObject
        $networkadapter = Get-CimInstance CIM_NetworkAdapter -Filter "DeviceID = '$($Adapter.index)'"
        $Output | Add-Member NoteProperty "PC_Name" $computername
        $output | Add-Member NoteProperty "Network_Adapter_Name" $networkadapter.Name
        # IPV4 address seems to always be first, would include IPV6 but the none loopback address index differs between adapters
        $output | Add-Member NoteProperty "Network_Adapter_IPV4_Address" $Adapter.Ipaddress[0]
        $output | Add-Member NoteProperty "Network_Adapter_AdapterType" $networkadapter.AdapterType
        $output | Add-Member NoteProperty "Network_Adapter_MACAddress" $networkadapter.MACAddress
        $output | Add-Member NoteProperty "Network_Adapter_Speed" ($networkadapter.Speed / 1Gb)
        $DGVNetInfo.Rows.Add(($Output).PC_Name, ($Output).Network_Adapter_Name, ($Output).Network_Adapter_IPV4_Address, ($Output).Network_Adapter_AdapterType, ($Output).Network_Adapter_MACAddress, ($Output).Network_Adapter_Speed)
    }
}

function Remove-Row {
    # Simple switch to remove selected row from DGV; event handler is on double click and is found under each DGV object
    # unfortunately double clicking deselectes other rows so you can't actually delete more than one, might change event handler in the future or you can just press delete on the kb
    switch ($tabControl1.SelectedIndex) {
        0 { $DGVPCInfo.SelectedRows | ForEach-Object { $DGVPCInfo.Rows.Remove($_) }; Break }
        1 { $DGVOSInfo.SelectedRows | ForEach-Object { $DGVOSInfo.Rows.Remove($_) }; Break }
        2 { $DGVCPUInfo.SelectedRows | ForEach-Object { $DGVCPUInfo.Rows.Remove($_) }; Break }
        3 { $DGVGPUInfo.SelectedRows | ForEach-Object { $DGVGPUInfo.Rows.Remove($_) }; Break }
        4 { $DGVNetInfo.SelectedRows | ForEach-Object { $DGVNetInfo.Rows.Remove($_) }; Break }
        Default { Break }
    }
}

function logAction ($action) {
    # Function for logging errors but can log other stuff too 
    If (-not(Test-Path "$ENV:USERPROFILE\Query-Computer Log.txt")) {
        New-Item "$ENV:USERPROFILE\Query-Computer Log.txt"
        Add-Content "$ENV:USERPROFILE\Query-Computer Log.txt" "$(Get-Date) | Log file has been created"
    }
    Add-Content "$ENV:USERPROFILE\Query-Computer Log.txt" "$(Get-Date) | $action"
}

Trap {
    # Catch-All error handling, should "dispose" of the form on error but I've never gotten a terminating error so not sure
    $wshell.Popup("An error occured! `nDetails can be found at $env:USERPROFILE\Query-Computer log.txt.", 0, "Crashed", 0x0)
    logAction "$($error | select-object -first 1)"
    $FMMain.Dispose()
}

$FMmain.ShowDialog()