# https://learn.microsoft.com/en-us/windows/win32/shell/shell-namespace
# https://learn.microsoft.com/en-us/windows/win32/shell/folder
# https://learn.microsoft.com/en-us/windows/win32/shell/folder-items
param([string]$phoneName = "Apple iPhone"
  , [string]$sourceFolder = "Internal Storage\DCIM"
  , [string]$targetFolder
  , [string]$filter = '(.jpg)|(.mp4)$')

function Get-ShellProxy {
  if ( -not $global:ShellProxy) {
    $global:ShellProxy = new-object -com Shell.Application
  }
  $global:ShellProxy
}

function Get-Phone {
  param($phoneName)
  $shell = Get-ShellProxy
  # 17 (0x11) = ssfDRIVES from the ShellSpecialFolderConstants (https://msdn.microsoft.com/en-us/library/windows/desktop/bb774096(v=vs.85).aspx)
  # => "My Computer" â€” the virtual folder that contains everything on the local computer: storage devices, printers, and Control Panel.
  # This folder can also contain mapped network drives.
  $shellItem = $shell.NameSpace(17).self
  $phone = $shellItem.GetFolder.items() | where { $_.name -eq $phoneName }
  return $phone
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

function Get-Dir {
  param($parent)
  $count = 0;
  $name = $parent.Name
  Write-Output $name
  foreach ($item in $parent.GetFolder.items() | Sort-Object Name) {
    if ($item.IsFolder) {
      Get-Dir $item
    }
    Write-Output $item.GetFolder.name
  }
}

function Main {
  $phoneFolderPath = $sourceFolder
  $phone = Get-Phone -phoneName $phoneName
  $folder = Get-SubFolder -parent $phone -path $phoneFolderPath
  Get-Dir $folder
  #Write-Verbose $folder.Type
}

Main

<#
#$items = @( $folder.GetFolder.items() | where { $_.Name -match $filter } )
$items = $folder.GetFolder.items()
if ($items) {
  $totalItems = $items.count
  Write-Verbose "Total items: $totalItems"
  if ($totalItems -gt 0) {
    $count = 0;
    foreach ($item in $items) {
      ++$count
      $fileName = $item.Name
      $percent = [int](($count / $totalItems) * 100)
      Write-Progress -Activity "Processing Files in $phoneName\$phoneFolderPath" `
        -status "Processing File ${count} / ${totalItems} (${percent}%)" `
        -CurrentOperation $fileName `
        -PercentComplete $percent
      Write-Verbose $fileName
    }
    Write-Verbose "Total items: $totalItems"
  }
}
#>