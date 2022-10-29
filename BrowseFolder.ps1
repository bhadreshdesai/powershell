[CmdletBinding()]param([string]$phoneName = "Downloads"
	, [string]$sourceFolder = "DCIM"
	, [string]$targetFolder = "C:\BDD\Backup")
	
function Get-ShellProxy {
	if ( -not $global:ShellProxy) {
		$global:ShellProxy = new-object -com Shell.Application
	}
	$global:ShellProxy
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

function HandleFile {
	param($file, $parentPath)
	#$destinationFolderPath = Join-Path -Path $targetFolder -ChildPath $parentPath
	#$destinationFolder = GetOrCreateFolder -folderPath $destinationFolderPath
	#$destinationFolder.GetFolder.CopyHere($file)
	Write-Verbose "HandleFile, destinationFolder: $destinationFolder"
	$filePath = Join-Path -Path $parentPath -ChildPath $file.Name
	Write-Verbose "HandleFile, filePath: $filePath"
}

function HandleFolder {
	param($folder, $parentPath)

	$destinationFolderPath = Join-Path $targetFolder $parentPath $folder.Name

	Write-Verbose "HandleFolder: $destinationFolderPath"

	#$currentPath = "$parentPath\$($folder.Name)"

	#$backupFolder = GetOrCreateFolder -destinationFolderPath $targetFolder\$currentPath
	#Write-Verbose $backupFolder

	# Write-Output $currentPath
}

function Get-FolderDetails {
	param($parent, $parentPath)
	$items = $parent.GetFolder.Items() | Sort-Object
	if ($items) {
		foreach ($item in $items) {
			if ($item.IsFolder()) {
				HandleFolder -folder $item -parentPath $parentPath
				Get-FolderDetails -parent $item -parentPath $parentPath
			}
			else {
				#$currentPath = "$parentPath\$($parent.Name)"
				$currentPath = Join-Path -Path $parentPath -ChildPath $parent.Name
				HandleFile -file $item -parentPath $currentPath
			}
		}
	}
}

function Get-SpecialFolder {
	param($specialFolderName)
	$shell = Get-ShellProxy
	# 17 (0x11) = ssfDRIVES from the ShellSpecialFolderConstants (https://msdn.microsoft.com/en-us/library/windows/desktop/bb774096(v=vs.85).aspx)
	# => "My Computer" â€” the virtual folder that contains everything on the local computer: storage devices, printers, and Control Panel.
	# This folder can also contain mapped network drives.
	$shellItem = $shell.NameSpace(17).self
	$specialFolder = $shellItem.GetFolder.items() | where { $_.name -eq $specialFolderName }
	return $specialFolder
}

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

function Main {
	$download = Get-SpecialFolder -specialFolderName $phoneName
	$dcim = Get-SubFolder -parent $download -path $sourceFolder
	# Write-Output "Download: $($download.Path); DCIM: $($dcim.Path)"
	Get-FolderDetails -parent $dcim -parentPath "DCIM"
	# Write-Host ($download | Format-Table | Out-String)
	# Write-Host ($dcim | Format-Table | Out-String)
}

Main