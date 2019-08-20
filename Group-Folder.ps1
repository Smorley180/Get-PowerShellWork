function Group-Folder {
<#
.Synopsis
Groups items found in sub-folders into parent folder.
.Description
Digs one folder deep and copies items into parent folder then removes old directories.
.Parameter Path
The path of the folder to be grouped.
.Parameter Force
If the force parameter is specified powershell will not prompt on directory deletion.
.Parameter Whatif
If the whatif parameter is specified powershell will display what would have happened in the host.
.Example
PS C:\> Group-Folder -Path C:\ -force
Will pull items out of subdirectories into C: drive; don't actually run this if Windows is on C:.
#>
    [cmdletBinding(SupportsShouldProcess)]
    param (
        # Path to folder to group.
        [Parameter(Mandatory = $true)]
        [string]
        $Path,
        # Skip confirmation on delete.
        [Parameter(Mandatory = $false)]
        [switch]
        $force
    )
    # List of directories in folder.
    $items = Get-ChildItem -Directory -Path $Path
    foreach ($d in $items) {
        # For each directory found move its childitems into the group folder and remove old directory.
        # If -whatif is specified displays info on host.
        if ($PSCmdlet.ShouldProcess($d.FullName, 'Copy')) {
            Get-ChildItem -Path $d.FullName | Copy-Item -Destination $Path -Recurse
        }
        if ($PSCmdlet.ShouldProcess($d.FullName, 'Remove')) {
            # If force is specified run the command without -confirm parameter.
            if ($force) {
                Remove-Item -Path $d.FullName -Force -Recurse
            }
            else {
                Remove-Item -Path $d.FullName -Force -Recurse -Confirm
            }
        }
    }
}
