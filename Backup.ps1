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

<#
Get Shell.Application com object
Reuse it by storing it on $global
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
    # 17 (0x11) = ssfDRIVES from the ShellSpecialFolderConstants (https://learn.microsoft.com/en-gb/windows/win32/api/shldisp/ne-shldisp-shellspecialfolderconstants?redirectedfrom=MSDN)
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
    $ssfDRIVES = 0x11
    return Get-FolderFromNamespace -namespace $ssfDRIVES -folderName $folderName
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

function GetOrCreateFolder {
    param($folderPath)

    # If destination path doesn't exist, create it only if we have some items to move
    if (-not (test-path -PathType Container $folderPath) ) {
        $folder = new-item -itemtype directory -path $folderPath
    }
    $shell = Get-ShellProxy
    $folder = $shell.Namespace($folderPath).self
    
    return $folder
}
function HandleFolder {
    param ($folder, $relativePath)
    # Get relative path from source folder
    # Write-Verbose $folder.Path
    if ($global:mode -eq "SIZE") {
        $global:totalFolderCnt ++
    }
    else {
        ++$global:folderCnt
        $percent = [int](($global:folderCnt * 100) / $global:totalFolderCnt)
        # Write-Verbose "Folder percentage: $percent" 
        $folderPath = $folder.Path
        $FolderProgressParameters = @{
            ID              = 1
            Activity        = 'Processing Folder'
            Status          = "Folder: ${folderPath} [${global:folderCnt} / ${global:totalFolderCnt} (${percent}%)]"
            PercentComplete = $percent
            #CurrentOperation = $folder.Path
        }
        Write-Progress @FolderProgressParameters
    }
}

function HandleFile {
    param ($file, $relativePath)
    #Write-Verbose $file.Path
    if ($global:mode -eq "SIZE") {
        $global:totalFileCnt ++
        $global:totalSize += $file.Size
    }
    else {
        $fileName = Split-Path $file.Path -Leaf
        $destinationFolderPath = Join-Path $backupFolderPath $relativePath
        # Check the target file doesn't exist:
        $targetFilePath = join-path -path $destinationFolderPath -childPath $fileName

        ++$global:fileCnt
        $global:size += $file.Size
        $percentFile = [int](($global:fileCnt * 100) / $global:totalFileCnt)
        # $percentSize = [int](($global:size * 100) / $global:totalSize)

        $filePath = $file.Path
        $FileProgressParameters = @{
            ID              = 2
            Activity        = 'Processing File'
            Status          = "File: ${filePath} [${global:fileCnt} / ${global:totalFileCnt} (${percentFile}%)]"
            PercentComplete = $percentFile
            #CurrentOperation = $folder.Path
        }
        Write-Progress @FileProgressParameters

        if (test-path -path $targetFilePath) {
            # write-error "Destination file exists - file not moved:`n`t$targetFilePath"
        }
        else {
            $destinationFolder = GetOrCreateFolder -folderPath $destinationFolderPath
            $destinationFolder.GetFolder.CopyHere($file)
            if (test-path -path $targetFilePath) {
                # Optionally do something with the file, such as modify the name (e.g. removed phone-added prefix, etc.)
            }
            else {
                write-error "Failed to move file to destination:`n`t$targetFilePath"
            }
        }

    }
}

function ProcessDir {
    param($folder, $parentPath)
    foreach ($item in $folder.GetFolder.items() | Sort-Object Name) {
        if ($item.IsFolder) {
            if ($parentPath) {
                $relativePath = Join-Path $parentPath $item.Name
            }
            else {
                $relativePath = $item.Name
            }
            HandleFolder -folder $item -relativePath $parentPath
            ProcessDir -folder $item -parentPath $relativePath
        }
        else {
            HandleFile -file $item -relativePath $parentPath
        }
    }
}

function CalculateSize {
    param($folder)
    $global:mode = "SIZE"
    ProcessDir -folder $folder
}

function CopyFiles {
    param($folder)
    $global:mode = "COPY"
    ProcessDir -folder $folder
}

function Init {
    $global:totalFolderCnt = 0
    $global:totalFileCnt = 0
    $global:totalSize = 0
    $global:folderCnt = 0
    $global:fileCnt = 0
    $global:size = 0
    $PSStyle.Progress.View = 'Classic'
}
function Main {
    Init
    $folderUnderMyComputer = Get-FolderUnderMyComputer -folderName $folderNameUnderMyComputer
    $sourceFolder = Get-SubFolder -parent $folderUnderMyComputer -path $sourceFolderPath
    
    CalculateSize -folder $sourceFolder

    Write-Verbose "FolderCnt: $global:totalFolderCnt, FileCnt: $global:totalFileCnt, Total Size: $global:totalSize"
    CopyFiles -folder $sourceFolder
    #Write-Verbose "FolderCnt: $global:totalFolderCnt, FileCnt: $global:totalFileCnt, Total Size: $global:totalSize"
    #Write-Verbose $sourceFolder.Path
}

Main