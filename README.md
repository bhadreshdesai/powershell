# Backup iPhone

## Instructions

Open powershell in the current folder
Execute the backup using the following command

```shell
. 'C:\BDD\powershell/Backup.ps1' -verbose -folderNameUnderMyComputer "OS (C:)" -sourceFolderPath "BDD\pstest\DCIM" -backupFolderPath "C:\BDD\pstest\Backup"
```

## TODO

- [x] Implement copy operation
- [ ] Tidy up code
- [ ] Implement delete operation
- [ ] Avoid using $global

## References

[How to reliably copy items with PowerShell to an mtp device?](https://stackoverflow.com/questions/55628092/how-to-reliably-copy-items-with-powershell-to-an-mtp-device)

[Folder.CopyHere method](https://learn.microsoft.com/en-us/windows/win32/shell/folder-copyhere)
