$here = Split-Path -parent $MyInvocation.MyCommand.Path
import-module (Join-Path (Split-Path -parent $here) '\poshTestEnvironment\poshTestEnvironment.ps1') -Force

Describe "Get-VoiceSynthesizer" {
    It "Returns voice info" {
        Get-VoiceSynthesizer | Should -BeOfType System.Speech.Synthesis.VoiceInfo
    }

    It "Can return current culture only." {
        $actual = (Get-VoiceSynthesizer -Culture).Culture
        $actual | Should -Not -Contain "en-US"
    }
}

Describe "New-VoiceSynthesizer" {
    It "Throws if given an unusable voice" {
        { New-VoiceSynthesizer -voice "This voice won't work" } | Should -Throw
    }

    It "Can create a new file" {
        New-VoiceSynthesizer -speech "" -file "$ENV:TEMP\temp.wav"
        "$ENV:TEMP\temp.wav" | Should -Exist
        Remove-Item "$ENV:TEMP\temp.wav"
    }

    It "Writes speech to file" {
        New-VoiceSynthesizer -Speech "One" -file "$ENV:TEMP\temp.wav" -speechrate 1 -voice "Microsoft David Desktop" 
        $file = (Get-Content "$ENV:TEMP\temp.wav")
        $file.length | Should -Be 194
        Remove-Item "$ENV:TEMP\temp.wav"
    }

    It "Can have different speech rates" {
        New-VoiceSynthesizer -Speech "One" -file "$ENV:TEMP\temp.wav" -speechrate 1 -voice "Microsoft David Desktop" 
        $file = (Get-Content "$ENV:TEMP\temp.wav")
        $file.length | Should -Be 194
        Remove-Item "$ENV:TEMP\temp.wav"

        New-VoiceSynthesizer -Speech "One" -file "$ENV:TEMP\temp.wav" -speechrate 10 -voice "Microsoft David Desktop" 
        $file = (Get-Content "$ENV:TEMP\temp.wav")
        $file.length | Should -Be 65
        Remove-Item "$ENV:TEMP\temp.wav"
    }
}

Describe "New-Rectangle" {
    It "Accepts only set colors" {
        { New-Rectangle -File New-TemporaryFile -colour "Black" -Force } | Should -Not -Throw
        { New-Rectangle -File New-TemporaryFile -colour "Yellow" -Force } | Should -Not -Throw
        { New-Rectangle -File New-TemporaryFile -colour "Red" -Force } | Should -Not -Throw
        { New-Rectangle -File New-TemporaryFile -colour "White" -Force } | Should -Not -Throw
        { New-Rectangle -File New-TemporaryFile -colour "Green" -Force } | Should -Not -Throw
        { New-Rectangle -File New-TemporaryFile -colour "Blue" -Force } | Should -Not -Throw
        { New-Rectangle -File New-TemporaryFile -colour "Orange" -Force } | Should -Not -Throw
        { New-Rectangle -File New-TemporaryFile -colour "Silver" -Force } | Should -Throw
    }

    It "Creates a file if none is found" {
        New-Rectangle -file "$ENV:TEMP\temp.jpg"
        "$ENV:TEMP\temp.jpg" | Should -Exist
        Remove-Item "$ENV:TEMP\temp.jpg"
    }

    It "Can have different heights" {
        New-Rectangle -file "$ENV:TEMP\temp.jpg" -height 200 -Colour "Blue"
        $File = (Get-Content "$ENV:TEMP\temp.jpg")
        $File.Length | Should -be 5
        Remove-Item "$ENV:TEMP\temp.jpg"

        New-Rectangle -file "$ENV:TEMP\temp.jpg" -height 400 -Colour "Blue"
        $File = (Get-Content "$ENV:TEMP\temp.jpg")
        $File.Length | Should -be 6
        Remove-Item "$ENV:TEMP\temp.jpg"
    }

    It "Can have different widths" {
        New-Rectangle -file "$ENV:TEMP\temp.jpg" -width 200 -Colour "Blue"
        $File = (Get-Content "$ENV:TEMP\temp.jpg")
        $File.Length | Should -be 4
        Remove-Item "$ENV:TEMP\temp.jpg"

        New-Rectangle -file "$ENV:TEMP\temp.jpg" -width 400 -Colour "Blue"
        $File = (Get-Content "$ENV:TEMP\temp.jpg")
        $File.Length | Should -be 5
        Remove-Item "$ENV:TEMP\temp.jpg"
    }

    Context "If file already exists" {
        Mock Read-Host { return "N" } -ParameterFilter { $Prompt -eq "$($File) already exists are you sure you want to overwrite? Y/N" }


        It "Prompts to overwrite" {
            New-Item "$env:TEMP\Test.jpg"
            $File = Get-Content "$env:TEMP\Test.jpg"
            $File.Length | Should -be 0
            New-Rectangle "$env:TEMP\Test.jpg"
            $File.Length | Should -be 0
            Remove-Item "$env:TEMP\Test.jpg"
        }
    }
    Context "If file already exists and you want to replace it" {
        Mock Read-Host { return "Y" } -ParameterFilter { $Prompt -eq "$($File) already exists are you sure you want to overwrite? Y/N" }

        It "Overwrites the file" {
            New-Item "$env:TEMP\Test.jpg"
            $File = Get-Content "$env:TEMP\Test.jpg"
            $File.Length | Should -be 0
            New-Rectangle "$env:TEMP\Test.jpg" -Colour "Blue"
            $File = Get-Content "$env:TEMP\Test.jpg"
            $File.Length | Should -be 5
            Remove-Item "$env:TEMP\Test.jpg"
        }
    }
}

Describe "New-LorumIpsum" {
    Context "When the website is up" {
        Mock Invoke-RestMethod { return "Lorum-Ipsum" } -ParameterFilter { $method -eq "GET" }
        It "Returns Lorum-Ipsum text" {
            New-LorumIpsum | Should -be "Lorum-Ipsum"
        }
    }
    
    Context "When the website is down" {
        Mock Invoke-RestMethod -MockWith { Throw }

        It "Returns an error" {
            New-LorumIpsum | Should -Be @'
An error occured when using Invoke-RestMethod: ScriptHalted 
'@
        }
    }
}

Describe "New-TestFiles" {
    Context "Only uses set extensions" {
        Mock New-Item { return $Null }
        Mock Get-Random { return 1 }

        It "Accepts only set extensions" {
            { New-TestFiles -FileType ".txt" -directory $ENV:TEMP } | Should -Not -Throw
            { New-TestFiles -FileType ".csv" -directory $ENV:TEMP } | Should -Not -Throw
            { New-TestFiles -FileType ".xls" -directory $ENV:TEMP } | Should -Not -Throw
            { New-TestFiles -FileType ".png" -directory $ENV:TEMP } | Should -Not -Throw
            { New-TestFiles -FileType ".wav" -directory $ENV:TEMP } | Should -Not -Throw
            { New-TestFiles -FileType ".jpeg" -directory $ENV:TEMP } | Should -Not -Throw
            { New-TestFiles -FileType ".tiff" -directory $ENV:TEMP } | Should -Not -Throw
            { New-TestFiles -FileType ".sql" -directory $ENV:TEMP } | Should -Throw
        }
    }
    Context "Makes random files in a directory" {

        It "Creates files" {
            $date = Get-Date
            $OriginalFiles = Get-ChildItem $env:TEMP
            New-TestFiles -directory $env:TEMP
            $NewFiles = Get-ChildItem $env:TEMP
            $OriginalFiles.Count | Should -BeLessThan $NewFiles.Count
            Get-Childitem $env:TEMP | Where-Object CreationTime -gt $date | Remove-Item
        }
    }

    Context "Can make files of a set type" {

        It "Makes files of a specific type" {
            New-TestFiles -directory $env:TEMP -FileType ".tiff"
            $NewFiles = Get-ChildItem $env:TEMP
            $NewFiles | Where-Object Extension -like "*.tiff" | Should -not -BeNullOrEmpty
            $NewFiles | Where-Object Extension -like "*.tiff" | Remove-Item
        }
    }
}

Describe "New-DirectoryStructure" {
    Context "Can create a random amount of folders" {
        Mock Get-Random { return 5 }
        # Still a bit slow but shouldn't make folders with same name breaking the test
        Mock New-TestingName { $Random = New-Object System.Random; Start-Sleep 0.2; $Rand = New-Object System.Random; Return "$($Random.Next())" + "$($Rand.Next())" }

        It "Creates folders" {
            New-DirectoryStructure -Directory "$env:TEMP\Test" -Depth 1
            $files = Get-ChildItem "$env:TEMP\Test"
            $files.count | Should -be 5
            $files | Remove-Item
        }
        
        It "Creates folders recursively" {
            New-DirectoryStructure -Directory "$env:TEMP\Test" -Depth 2
            $files = Get-ChildItem "$env:TEMP\Test" -Recurse
            $files.Count | Should -be 30
            # Slightly overkill but this way there are no prompts or errors to slow things down
            $files | Remove-Item -Force -Recurse -ErrorAction "SilentlyContinue"
        }
    }
}