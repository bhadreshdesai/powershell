[CmdletBinding()]
param (
    [Parameter(Position = 0, mandatory = $true)]    
    [string] $folderNameUnderMyComputer,
    [Parameter(Position = 1, mandatory = $true)]
    [string] $sourceFolderPath,
    [Parameter(Position = 2, mandatory = $true)]
    [string] $backupFolderPath
)

<#
Backup.ps1 -folderNameUnderMyComputer "OS (C:)" -sourceFolderPath "BDD\pstest\DCIM" -backupFolderPath "C:\BDD\pstest\Backup"
#>

function Get-ShellProxy {
    if ( -not $global:ShellProxy) {
        $global:ShellProxy = new-object -com Shell.Application
    }
    $global:ShellProxy
}

<#
Get a folder from a given namespace
#>
function Get-FolderFromNamespace {
    param($namespace, $folderName)
    $shell = Get-ShellProxy
    # 17 (0x11) = ssfDRIVES from the ShellSpecialFolderConstants (https://msdn.microsoft.com/en-us/library/windows/desktop/bb774096(v=vs.85).aspx)
    # => "My Computer" â€” the virtual folder that contains everything on the local computer: storage devices, printers, and Control Panel.
    # This folder can also contain mapped network drives.
    # $shellItem = $shell.NameSpace(17).self
    # 17 This PC, 5 Documents 
    # $shellItem = $shell.NameSpace("shell:$($specialFolderName)").self
    $shellItem = $shell.NameSpace($namespace).self
    $folder = $shellItem.GetFolder.items() | where { $_.name -eq $folderName }
    return $folder
}

<#
Get folder under My Computer
#>
function Get-FolderUnderMyComputer {
    param (
        $folderName
    )
    return Get-FolderFromNamespace -namespace 17 -folderName $folderName
}

<#
Get subfolder relative to the parent
#>
function Get-SubFolder {
    param($parent, [string]$path)
    $pathParts = @( $path.Split([system.io.path]::DirectorySeparatorChar) )
    $current = $parent
    foreach ($pathPart in $pathParts) {
        if ($pathPart) {
            $current = $current.GetFolder.items() | where { $_.Name -eq $pathPart }
        }
    }
    return $current
}

function HandleFolder {
    param ($folder)
    # Get relative path from source folder
    Write-Verbose $folder.Path
}

function HandleFile {
    param ($file)
    Write-Verbose $file.Path
}

function Get-Dir {
    param($parent)
    foreach ($item in $parent.GetFolder.items() | Sort-Object Name) {
        if ($item.IsFolder) {
            HandleFolder -folder $item
            Get-Dir $item
        }
        else {
            HandleFile -file $item
        }
    }
}
function Main {
    $folderUnderMyComputer = Get-FolderUnderMyComputer -folderName $folderNameUnderMyComputer
    $sourceFolder = Get-SubFolder -parent $folderUnderMyComputer -path $sourceFolderPath
    Get-Dir -parent $sourceFolder
    #Write-Verbose $sourceFolder.Path
}

Main