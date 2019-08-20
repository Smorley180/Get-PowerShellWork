<#
.Synopsis
Copy folder structure and compress bottom level folders.
.Description
This command will compress bottom level folders using 7zip to either .7z or .zip and copy them
and the folder structure to a seperate folder.
.Parameter Path
The path of the folder to be backed-up. By default it will use the folder you are currently in.
.Parameter Destination
The path to the destination folder.
.Parameter FileType
The type of file to be compressed to, accepts 7z or zip. Uses 7z by default.
.Parameter Compression
Compression level, defaults to 5. Accepts 1, 3, 5, 7, 9 where higher is more compressed.
For more information see https://sevenzip.osdn.jp/chm/cmdline/switches/method.htm
.Example
PS C:\> .\Backup-Folder C:\users\test C:\users\test2

Backup folder \test to \test2

.Example
PS C:\> .\Backup-Folder -destination $env:ProgramFiles\Backups -FileType zip -Compression 9

Backups current folder to ProgramFiles\Backups, uses zip compression at level 9.

#>
[Cmdletbinding()]
param (
    $path = (Get-Location),
    [Parameter(Mandatory = $True)]
    $destination,
    [ValidateSet("7z", "zip")]
    $FileType = "7z",
    [ValidateSet("1", "3", "5", "7", "9")]
    $Compression = 5
)

$ErrorActionPreference = "Stop"

#Confirm 7zip is installed; throws if not. Sets an alias for 7zip.

if (-not (test-path "$env:ProgramFiles\7-Zip\7z.exe")) { throw "$env:ProgramFiles\7-Zip\7z.exe needed" }
set-alias sz "$env:ProgramFiles\7-Zip\7z.exe"

#Grab the last item in the file path for later.

$SplitName = ($path | split-path -leaf)

<#Loop through each item in $dir, using split-path to get the parent folder then take two steps: If the parent folder
doesn't exist in $children add the full path to $children, or if the parent folder does exist in $children it must
have child folders so remove it and add the full path.
#>
Write-Verbose "Getting files to compress."
$dirs = Get-ChildItem ($path) -Recurse -Directory
$children = New-Object System.Collections.ArrayList
ForEach ($d in $dirs) {
    $r = split-path $d.FullName -Parent
    if ($children -notcontains $r) {
        $children.add($d.fullname) | Out-Null
    }
    else {
        $children.remove($r)
        $children.add($d.fullname) | Out-Null
    }
}

Write-Verbose "Compressing files."
ForEach ($d in $children) {
    sz a -t"$FileType" -mx"$Compression" "$d.$FileType" "$d" | Out-Null
}

# Find the source files

Write-Verbose "Copying files."
Get-ChildItem $path -Recurse *.$FileType | ForEach-Object {

    # Remove the root folders.

    $split = $_.Fullname -split "$SplitName\\"
    $DestFile = $split[1]

    # Build the new destination file path.

    $DestFile = "$destination\$Destfile"

    <# Copy-Item won't create the folder structure so we have to
     create a blank file and then overwrite it. #>
     
    $null = New-Item -Path  $DestFile -Type File -Force
    Move-Item -Path  $_.FullName -Destination $Destfile -Force
}
Write-Verbose "Folder successfully backed up."