Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$FMMain = New-Object System.Windows.Forms.Form

$DGVOutput = New-Object System.Windows.Forms.DataGridView
$TBName = New-Object System.Windows.Forms.TextBox
$Label = New-Object System.Windows.Forms.Label
$BTGetInfo = New-Object System.Windows.Forms.Button
$BTClear = New-Object System.Windows.Forms.Button
#
# DGVOutput
#
$DGVOutput.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::AutoSize
$DGVOutput.Location = New-Object System.Drawing.Point(13, 64)
$DGVOutput.Name = "DGVOutput"
$DGVOutput.Size = New-Object System.Drawing.Size(698, 359)
$DGVOutput.TabIndex = 3
$DGVOutPut.RowHeadersVisible = $false
$DGVOutPut.AutoSizeColumnsMode = 'Fill'
$DGVOutPut.AllowUserToResizeRows = $false
$DGVOutPut.selectionmode = 'FullRowSelect'
$DGVOutPut.MultiSelect = $false
$DGVOutPut.AllowUserToAddRows = $false
$DGVOutPut.ReadOnly = $true
$DGVOutPut.ColumnCount = 5
$DGVOutPut.ColumnHeadersVisible = $true
$DGVOutput.Columns[0].Name = "PC_Name"
$DGVOutPut.Columns[0].Width = 100
$DGVOutPut.Columns[1].Name = "OS_Type"
$DGVOutPut.Columns[2].Name = "OS_Version"
$DGVOutPut.Columns[2].Width = 70
$DGVOutPut.Columns[3].Name = "OS_SystemDrive"
$DGVOutPut.Columns[3].Width = 90
$DGVOutPut.Columns[4].Name = "OS_SerialNumber"
$DGVOutPut.Columns[4].Width = 160
$DGVOutput.add_doubleclick( {
	Copy-Info
})
#
# TBName
#
$TBName.Location = New-Object System.Drawing.Point(60, 24)
$TBName.Name = "TBName"
$TBName.Size = New-Object System.Drawing.Size(217, 20)
$TBName.TabIndex = 0
#
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
# BTGetInfo
#
$BTGetInfo.Location = New-Object System.Drawing.Point(340, 24)
$BTGetInfo.Name = "BTGetInfo"
$BTGetInfo.Size = New-Object System.Drawing.Size(75, 20)
$BTGetInfo.TabIndex = 1
$BTGetInfo.Text = "Get Info"
$BTGetInfo.UseVisualStyleBackColor = $true
$BTGetInfo.add_click( {
        Get-OSInfo
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
        $DGVOutput.Rows.Clear()
    })
#
# FMMain
#
$FMMain.ClientSize = New-Object System.Drawing.Size(739, 434)
$FMMain.Controls.Add($BTClear)
$FMMain.Controls.Add($BTGetInfo)
$FMMain.Controls.Add($Label)
$FMMain.Controls.Add($TBName)
$FMMain.Controls.Add($DGVOutput)
$FMMain.Name = "FMMain"
$FMMain.Text = "Get-OSInfoGUI"

function Get-OSInfo {
    $Output = New-Object -TypeName PSObject
    $OS = Get-CIMInstance CIM_operatingsystem -computername $TBName.Text
    $Output | Add-Member ScriptProperty "PC_Name" { if ($TBName.Text -gt 0) {
            return $TBName.Text
        }
        else {
            return $env:COMPUTERNAME
        } }
    $Output | Add-Member NoteProperty "OS_Type" $OS.Caption
    $Output | Add-Member NoteProperty "OS_Version" $OS.Version
    $Output | Add-Member NoteProperty "OS_SystemDrive" $OS.SystemDrive
    $Output | Add-Member NoteProperty "OS_SerialNumber" $OS.SerialNumber
    $DGVOutput.Rows.Add(($Output).PC_Name, ($Output).OS_Type, ($Output).OS_Version, ($Output).OS_SystemDrive, ($Output).OS_SerialNumber)
    $TBName.Clear()
}

function Copy-Info {
    Set-Clipboard $DGVOutput.SelectedCells[4].Value
}
$FMmain.ShowDialog()
