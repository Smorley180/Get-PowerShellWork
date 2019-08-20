<#class VoiceNames : System.Management.Automation.IValidateSetValuesGenerator {
    [String[]] GetValidValues() {
        Add-Type -AssemblyName System.Speech
        $Speech = New-Object System.Speech.Synthesis.SpeechSynthesizer
        $Voices = $Speech | Select-Object -ExpandProperty VoiceInfo
        return [String[]] $Voices
    }
    
}#>


function Get-VoiceSynthesizer {
    <#
.Synopsis
Gets a list of currently enabled voice synthesizers.
.Description
Pulls a list of voice synthesizers currently set to Enabled, not all voice synthesizers are shown here if they have been set to disabled in the registry.
.Parameter Culture
Gets only the voices for the currently used culture.
.Example
PS C:\> Get-VoiceSynthesizer
Gets the currently enabled voice synthesizers.
.Example
PS C:\> Get-VoiceSynthesizer -Culture
Gets currently enabled voice synthesizers that belong to the currently used culture.
#>
    param (
        # Gets only the voices for the current culture
        [Parameter(Mandatory = $false)]
        [Switch]
        $Culture
    )
    # Add the needed assembly.
    Add-Type -AssemblyName System.Speech
    # Create the speech synthesizer
    $Speech = New-Object System.Speech.Synthesis.SpeechSynthesizer
    switch ($Culture) {
        # Get installed voices.
        $True { $Voices = $Speech.GetInstalledVoices($(Get-Culture)); Break }
        $False { $Voices = $Speech.GetInstalledVoices(); Break }
        Default { Break }
    }
    # Return the voices.
    return $Voices | Select-Object -ExpandProperty VoiceInfo
}

function New-VoiceSynthesizer {
    <#
.Synopsis
Creates text-to-speech out of user input.
.Description
Uses Windows Speech Synthesizer and user input to create text-to-speech, can be saved to a .wav file.
.Parameter File
The file to save the text-to-speech to.
.Parameter SpeechRate
How fast the synthesizer will speak, accepts a range between 1 and 10 and defaults to 1.
.Parameter Speech
The string to speak, defaults to using the New-LorumIpsum command.
.Parameter Voice
The voice to use; the voices available and the default voice depend on installed language packs. Run Get-VoiceSynthesizers to see a list of all currently available voices.
.Example
PS C:\> New-VoiceSynthesizer
Creates a random text-to-speech.
.Example
PS C:\> New-VoiceSynthesizer -SpeechRate 3 -Speech "Robot voice." -Voice "Microsoft David Desktop"
Speaks the phrase "Robot Voice" with a speed of 3 and through the tones of David.
.Example
PS C:\> New-VoiceSynthesizer -file "F:\Phrases\phrase.wav"
Creates a random text-to-speech and saves it to "F:\Phrases\phrase.wav".
#>
    [cmdletbinding()]
    param (
        # .wav File to save to
        [Parameter(Mandatory = $false,
            ValueFromPipeline = $true)]
        [String]
        $File,
        # The speed at which the voice talks.
        [Parameter(Mandatory = $false)]
        [ValidateRange(-10, 10)]
        [int]
        $SpeechRate = 1,
        # Text to Speak
        [Parameter(Mandatory = $false)]
        [String]
        $Speech = $(New-LorumIpsum -TextLength Short -TextAmount 1),
        # Voice to use
        [Parameter(Mandatory = $false)]
        [String]
        $Voice
    )
    Begin {
        # Add the needed assembly.
        Add-Type -AssemblyName System.Speech
        # Create the Speech Synthesizer
        $SpeechSynthesizer = New-Object System.Speech.Synthesis.SpeechSynthesizer
        # Set the rate of the speech synthesizer.
        $SpeechSynthesizer.Rate = $SpeechRate
        $Voices = Get-VoiceSynthesizer
        if ($Voice) {
            # If the chosen voice exists on the system set the synthesizer to use it.
            if ($Voices.Name -contains $Voice) {
                $SpeechSynthesizer.SelectVoice($Voice)
            }
            else {
                # If it doesn't exist throw with an error.
                throw "$($Voice) is not currently enabled on this system, run Get-VoiceSynthesizers to see a list of all currently enabled voice synthesizers."
            }
        }
    }
    Process {
        # If a save file is chosen output to it.
        if ($File) {
            # Create the file if it doesn't exist
            if (-not(Test-Path $File)) {
                New-Item $File
            }
            $SpeechSynthesizer.SetOutputToWaveFile($File)
        }
        # Or output to audio.
        else {
            $SpeechSynthesizer.SetOutputToDefaultAudioDevice()
        }

        $SpeechSynthesizer.Speak($Speech)
    }
    End {
        $SpeechSynthesizer.Dispose()
    }

}

function New-Rectangle {
    <#
.Synopsis
Creates a rectangle of determinate color and size.
.Description
Uses windows drawing to create a bitmap and draws a rectangle on it.
.Parameter File
The file to save the image to.
.Parameter Width
Width of rectangle to create.
.Parameter Height
Height of rectangle to create.
.Parameter Colour
Colour of rectangle to creatre.
.Example
PS C:\> New-Rectangle
Creates a random text-to-speech.
.Example
PS C:\> New-VoiceSynthesizer -SpeechRate 3 -Speech "Robot voice." -Voice "Microsoft David Desktop"
Speaks the phrase "Robot Voice" with a speed of 3 and through the tones of David.
.Example
PS C:\> New-VoiceSynthesizer -file "F:\Phrases\phrase.wav"
Creates a random text-to-speech and saves it as phrase. 
#>
    param (
        # Location to save
        [Parameter(Mandatory = $true)]
        [string]
        $File,
        # Width of image
        [Parameter(Mandatory = $false)]
        [int]
        $Width = 250,
        # Height of image
        [Parameter(Mandatory = $false)]
        [int]
        $Height = 61,
        # Colour to use
        [Parameter(Mandatory = $false)]
        [ValidateSet("Black", "Yellow", "Red", "White", "Green", "Blue", "Orange")]
        [String]
        $Colour = (Get-Random("Black", "Yellow", "Red", "White", "Green", "Blue", "Orange"))
    )
    # Add the needed assembly.
    Add-Type -AssemblyName System.Drawing
    # Create the bitmap
    $bmp = New-Object System.Drawing.Bitmap $width, $Height
    # Assign the colour of brush to use.
    switch ($Colour) {
        "Black" { $brush = [System.Drawing.Brushes]::Black; Break }
        "Yellow" { $brush = [System.Drawing.Brushes]::Yellow; Break }
        "Red" { $brush = [System.Drawing.Brushes]::Red; Break }
        "White" { $brush = [System.Drawing.Brushes]::White; Break }
        "Green" { $brush = [System.Drawing.Brushes]::Green; Break }
        "Blue" { $brush = [System.Drawing.Brushes]::Blue; Break }
        "Orange" { $brush = [System.Drawing.Brushes]::Orange; Break }
        Default { Break }
    }
    $graphics = [System.Drawing.Graphics]::FromImage($bmp)
    # Create a rectangle 
    $graphics.FillRectangle($brush, 0, 0, $bmp.Width, $bmp.Height)
    $graphics.Dispose()
    # Save the bitmap to file.
    if (-not(Test-Path $file)) {
        New-Item $File
    }
    $bmp.Save($File)
    #return $bmp
}

function New-LorumIpsum {
    <#
.Synopsis
Creates Lorum Ipsum text.
.Description
Uses REST API to retrieve generated Lorum Ipsum from loripsum.net, thanks to them for making the generator.
.Parameter TextLength
The Length of the paragraphs to be generated.
.Parameter TextAmount
The amount of paragraphs to be generated.
.Parameter Headers
Adds headers to generated text.
.Parameter UnorderedLists
Adds lists to generated text.
.Example
PS C:\> $text = New-LorumIpsum -TextLength "Long" -TextAmount 30
Assigns lorum ipsum to a text variable.
.Example
PS C:\> New-LorumIpsum -Headers -UnorderedLists | Out-File .\LorumIpsum.txt
Exports lorum ipsum to a text file.
#>
    param (
        # Average length of paragraph to generate
        [Parameter(Mandatory = $false)]
        [ValidateSet("Short", "Medium", "Long", "Very Long")]
        [string]
        $TextLength,
        # Number of paragraphs to generate.
        [Parameter(Mandatory = $false)]
        [int]
        $TextAmount,
        # Add headers to text
        [Parameter(Mandatory = $false)]
        [switch]
        $Headers,
        # Add unordered lists to text
        [Parameter(Mandatory = $false)]
        [switch]
        $UnorderedLists
    )
    # Defines the site to get text from.
    $text = 'https://loripsum.net/api/plaintext/'
    # Switch to add the length of text onto the URL.
    switch ($TextLength) {
        "Short" { $text + "short/" | Out-Null; Break }
        "Medium" { $text + "medium/" | Out-Null; Break }
        "Long" { $text + "long/" | Out-Null; Break }
        "Very Long" { $text + "verylong/" | Out-Null; Break }    
        Default { Break }
    }
    # Add headers to generated text
    if ($Headers) {
        $text = $text + "ul/"
    }
    # Add lists to generated text
    if ($UnorderedLists) {
        $text = $text + "headers/"
    }
    # Add the amount of paragraphs onto the URL.
    $text = $text + "$($TextAmount)"
    # Return the generated text from REST method.
    try {
        return (Invoke-RestMethod -Method GET $text)
    }
    catch {
        "An error occured when using Invoke-RestMethod: $_ "
    }
}

function New-TestTable {
    <#
.Synopsis
Creates a table of random data.
.Description
Creates a table of random data for use in databases/AD.
.Parameter TableRows
The amount of objects to create in the table, by default it will create between 5 and 15.
.Parameter Product
Create a table with information about products.
.Parameter Employee
Create a table with information about employees.
.Example
PS C:\> New-TestFiles -directory "C:\Test"
Creates files.
#>
    [cmdletbinding(DefaultParameterSetName = "Employee")]
    param (
        # Amount of objects to generate
        [Parameter(Mandatory = $false)]
        [int]
        $TableRows = (Get-Random(5..15)),
        # Generate a product description table
        [Parameter(Mandatory = $false,
            ParameterSetName = "Product")]
        [switch]
        $Product,
        # Generate a Employee description table
        [Parameter(Mandatory = $false,
            ParameterSetName = "Employee")]
        [switch]
        $Employee
    )
    Begin {
        $Templates = Join-Path (Split-Path -Parent -Path $PSScriptRoot) '/templates'
    }
    Process {
        # If no table type is specified choose one at random
        if (-not($Employee) -and -not($Product)) {
            switch (Get-Random(1..2)) {
                1 { $Employee = $true; Break }
                2 { $Product = $true; Break }
                Default { Break }
            }
        }
        # Create our array to hold the generated items.
        $Table = @()
        if ($Employee) {
            # Load in the needed data files.
            $MaleNames = Get-Content "$templates\First_Names_M.txt"
            $FemaleNames = Get-Content "$templates\First_Names_F.txt"
            $LastNames = Get-Content "$templates\Last_Names.txt"
            $JobTitle = Get-Content "$templates\Job_Title.txt"
            # Create one data point for every row specified
            for ($i = 0; $i -lt $TableRows; $i++) {
                $Gender = Get-Random ("Male", "Female")
                $User = @{
                    "First Name" = If ($Gender -eq "Male") { Get-Random($MaleNames) } else { Get-Random($FemaleNames) }
                    "Last Name" = Get-Random($LastNames)
                    "Job Title" = Get-Random($JobTitle)
                }
                [PSCustomObject]$User
                # Add the data point to the array
                $Table += $User
            }
        }
        if ($Product) {
            $Prefix = Get-Content "$templates\Adjectives.txt"
            $Suffix = Get-Content "$templates\Nouns.txt"
            for ($i = 0; $i -lt $tablerows; $i++) {
                $User = @{
                    "Item" = "$(Get-Random($Prefix)) $(Get-Random($Suffix))"
                    "Cost" = "`$$(Get-Random(0..1000)).$(([string](Get-Random(0..99))).PadLeft(2,'0'))"
                    "Stock Remaining" = "$(Get-Random(1..1000))"
                }
                [PSCustomObject]$Item
                $Table += $Item
            }
        }
        Return $Table
    }
}

function New-TestFiles {
    <#
.Synopsis
Creates random files in a directory.
.Description
Creates a random amount of files and assigns each a random filetype.
.Parameter Directory
The directory to create the files in.
.Parameter FileType
Create files of one specific filetype instead of assorted filetypes.
.Example
PS C:\> New-TestFiles -directory "C:\Test"
Creates files.
#>
    param (
        # Directory to create files in.
        [parameter(ValueFromPipeline = $true)]
        [string]
        $directory,
        # Type of file to create
        [Parameter(Mandatory = $false)]
        [ValidateSet(".txt", ".csv", ".xls", ".png", ".wav", ".jpeg", ".tiff")]
        [string]
        $FileType
    )
    # If no filetype is specified create random ones.
    if (-not($FileType)) {
        # Create an amount of files between 5 and 20.
        for ($i = 0; $i -lt (Get-Random(5..20)); $i++) {
            # Randomly assign the type of file.
            $type = Get-Random(1..7)
            switch ($type) {
                1 { $type = ".txt"; Break }
                2 { $type = ".csv"; Break }
                3 { $type = ".xls"; Break }
                4 { $type = ".png"; Break }
                5 { $type = ".wav"; Break }
                6 { $type = ".jpeg"; Break }
                7 { $type = ".tiff"; Break }
                Default { Break }
            }
            # Create each new file using New-TestingName to create a random name.
            $filename = "$($directory)\$(New-TestingName)$($type)"
            New-Item -Path $filename
        }
    }
    # Else do the same thing but with the specified filetype.
    else {
        $type = $FileType
        for ($i = 0; $i -lt (Get-Random(5..20)); $i++) {
            $filename = "$($directory)\$(New-TestingName)$($type)"
            New-Item -Path $filename
        }
    }
}

function New-DirectoryStructure {
    <#
.Synopsis
Fills a directory up with recursive folders.
.Description
Creates folders inside folders up to a specified depth.
.Parameter Directory
The directory to create the folders in.
.Parameter Depth
The depth of the folder structure to create. By default it will go between 1 and 3 folders deep. Will see exponential slowdowns the higher you set this.
.Example
PS C:\> New-DirectoryStructure -directory "C:\Test" -depth 2
Creates folders with a depth of 2.
#>
    
    param (
        # Directory to create folders in
        [Parameter(Mandatory = $true)]
        [string]
        $Directory,
        # Depth of folder structures to create
        [Parameter(Mandatory = $false)]
        [int]
        $Depth = (Get-Random(1..3))
    )
    # Create one layer of folders for each "depth"
    for ($i = 0; $i -lt $Depth; $i++) {
        $dirs = Get-ChildItem $($Directory) -Recurse -Directory
        #Create an arraylist to hold our bottom level directories.
        $children = New-Object System.Collections.ArrayList
        # If there are no subfolders this will create the initial 'layer' of subfolders.
        if ($dirs.Count -lt 1) {
            $children.Add($Directory)
        }
        # Else if there are subfolders find the bottom level ones.
        else {
            ForEach ($d in $dirs) {
                # I pulled this from part of an old script which I took from the internet not super sure how it works but it is very fast 
                $r = split-path $d.FullName -Parent
                if ($children -notcontains $r) {
                    $children.add($d.fullname) | Out-Null
                }
                else {
                    $children.remove($r)
                    $children.add($d.fullname) | Out-Null
                }
            }
        }
        # For every directory in our arraylist create between 1 and 10 folders.
        foreach ($d in $children) {
            for ($j = 0; $j -lt (Get-Random(1..10)); $j++) {
                New-Item -Path $d -ItemType Directory -Name $(New-TestingName)
            } 
        }
    }
}

function New-TestingName {
    # Choose a number between 1 and 20
    $rca = (1..(Get-Random(2..20))) | ForEach-Object { $ran = Get-Random -Minimum 97 -Maximum 123; [char][byte]$ran } # For each number create a character
    $rca -join '' 
}

function New-TestDirectory {
    <#
.Synopsis
Combines the New-DirectoryStructure and New-TestFiles cmdlets.
.Description
Uses New-DirectoryStructure and New-TestFiles to fill a directory with folders and each folder with random files.
.Parameter Directory
The directory to create the folders in.
.Parameter Depth
The depth of the folder structure to create. By default it will go between 1 and 3 folders deep. Will see exponential slowdowns the higher you set this.
.Example
PS C:\> New-DirectoryStructure -directory "C:\Test" -depth 2
Creates folders with a depth of 2.
#>
    param (
        # Directory to populate
        [Parameter(Mandatory = $true)]
        [string]
        $directory,
        # Depth of folder structure to create.
        [Parameter(Mandatory = $false)]
        [int]
        $depth = (Get-Random(1..3))
    )
    # Create the directory structure
    New-DirectoryStructure -directory $directory -depth $depth
    # Create junk files in the top directory
    New-TestFiles -directory $directory
    $created = Get-ChildItem -Path $directory -Recurse -Directory
    # Add junk files to each directory that was created
    foreach ($create in $created) {
        New-TestFiles -directory $create.FullName
    }
}

function Set-TestFiles {
    <#
.Synopsis
Files files in a directory with testing information.
.Description
Utilises the New-LorumIpsum, New-Rectangle, New-VoiceSynthesizer and New-TestTable cmdlets to create test information and adds that to files.
.Parameter Directory
The directory of files that you want to add content to.
.Parameter TextAmount
The amount of paragraphs to be generated in text files.
.Parameter TextLength
The Length of the paragraphs to be generated in text files.
.Parameter TableRows
The amount of objects to create in the table, by default it will create between 4 and 8. Used for .csv and .xls.
.Example
PS C:\> $text = New-LorumIpsum -TextLength "Long" -TextAmount 30
Assigns lorum ipsum to a text variable.
.Example
PS C:\> New-LorumIpsum -Headers -UnorderedLists | Out-File .\LorumIpsum.txt
Exports lorum ipsum to a text file.
#>
    param (
        # Parameter help description
        [Parameter(Mandatory = $true)]
        [string]
        $Directory,
        # Set files in subdirectories
        [Parameter(Mandatory = $false)]
        [Switch]
        $Recurse,
        # Average length of paragraph to generate
        [Parameter(Mandatory = $false)]
        [ValidateSet("Short", "Medium", "Long", "Very Long")]
        [string]
        $TextLength = "Medium",
        # Number of paragraphs to generate.
        [Parameter(Mandatory = $false)]
        [int]
        $TextAmount = 10,
        # Amount of objects to generate
        [Parameter(Mandatory = $false)]
        [int]
        $TableRows = (Get-Random(4..8))
    )
    switch ($Recurse) {
        $True { $file_list = Get-ChildItem $Directory -Recurse; Break }
        $False { $file_list = Get-ChildItem $Directory; Break }
        Default { Break }
    }
    foreach ($file in ($file_list | Where-Object { $_.Extension -like ".csv" -or $_.Extension -like ".xls" })) {
        New-TestTable -TableRows $TableRows | Export-Csv -path $file.FullName -NoTypeInformation
    }
    foreach ($file in ($file_list | Where-Object { $_.Extension -like ".txt" } )) {
        Add-Content -path $file.FullName -Value (
            New-LorumIpsum -TextLength $TextLength -TextAmount $TextAmount)
    }
    foreach ($file in ($file_list | Where-Object { $_.Extension -like ".png" `
                    -or $_.Extension -like ".jpeg" -or $_.Extension -like ".tiff" })) {
        New-Rectangle -File $file.FullName
    }
    foreach ($file in ($file_list | Where-Object { $_.Extension -like ".wav" })) {
        New-VoiceSynthesizer -File $file.FullName
    }
}

function New-TestEnvironment {
    <#
.Synopsis
Combines the New-TestDirectory and Set-TestFiles commands to create a complete test environment.
.Description
Utilises the New-LorumIpsum, New-Rectangle, New-VoiceSynthesizer and New-TestTable cmdlets to create test information and adds that to files.
This cmdlet splits the path and the name of the test environment so the command isn't accidnelty run in a directory where it could cause damage.
.Parameter Path
Path to the location to create the test environment in. Defaults to the users home directory.
.Parameter Name
The name of the Test Environment to create. 
.Parameter TextLength
The Length of the paragraphs to be generated in text files.
.Parameter TextAmount
The amount of paragraphs to be generated in text files.
.Parameter TableRows
The amount of objects to create in the table, by default it will create between 4 and 8. Used for .csv and .xls.
.Parameter Depth
The amount of subfolders to create, randomizes between 1 and 3 if no value is given.
.Example
PS C:\> $text = New-LorumIpsum -TextLength "Long" -TextAmount 30
Assigns lorum ipsum to a text variable.
.Example
PS C:\> New-LorumIpsum -Headers -UnorderedLists | Out-File .\LorumIpsum.txt
Exports lorum ipsum to a text file.
#>
    param (
        # Specifies a folder location to generate the test directory.
        [Parameter(Mandatory = $false)]
        [String]
        $Path,
        # Name of the test directory to be created.
        [Parameter(Mandatory = $true)]
        [string]
        $Name,
        # Average length of paragraph to generate
        [Parameter(Mandatory = $false)]
        [ValidateSet("Short", "Medium", "Long", "Very Long")]
        [string]
        $TextLength,
        # Number of paragraphs to generate.
        [Parameter(Mandatory = $false)]
        [int]
        $TextAmount,
        # Depth of folder structure to create.
        [Parameter(Mandatory = $false)]
        [int]
        $depth = (Get-Random(1..3))
    )
    if (-not([String]::IsNullOrEmpty($Path))) {
        $Directory = $Path
    }
    else {
        $Directory = "$env:HOMEDRIVE\$env:HOMEPATH"
    }
    if (-Not(Test-Path $Directory\$Name)) {
        New-Item -Path $Directory -Name $Name -ItemType Directory
    }
    New-TestDirectory -directory "$($Directory)\$($Name)" -depth $depth
    Set-TestFiles -Directory "$($Directory)\$($Name)" -TextLength $TextLength -TextAmount $TextAmount -Recurse
}

#New-TestEnvironment -Path "F:\" -name "Testicles" -TextLength "Short" -TextAmount "1" -depth 2
#New-DirectoryStructure -directory "F:\Test" -depth 0
#New-LorumIpsum -TextLength "Very Long" -TextAmount 10 -Headers
#New-TestFiles -directory F:\Test2 -FileType .csv
#New-TestTable -TableRows 100 -Product
#New-TestTable | Export-Csv F:\Testicles\wvufdjcshjohlxae.csv -NoTypeInformation
#New-Rectangle -File F:\Testicles\nhucglbxogrxogqzs.tiff
#New-VoiceSynthesizer <#-File F:\Testicles\Speech.wav#> -Voice "Microsoft Hazel Desktop" -SpeechRate -10 -Speech "Hilarious banter mate."
#Get-VoiceSynthesizers