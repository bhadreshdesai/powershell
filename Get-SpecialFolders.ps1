[CmdletBinding()] param ()

$rtn = New-Object psobject -Property @{"Value" = $null; "name" = $null; "path" = $null }
1..500 |
ForEach-Object {
    $rtn.Value = $_
    $rtn.name = (new-object -com shell.application ).namespace($_).title
    $rtn.path = (new-object -com shell.application ).namespace($_).self.path
    if ($rtn.name -ne $null) {
        # “$($rtn.value) $($rtn.name) $($rtn.path)” | Out-File -FilePath $file -Append
        Write-Verbose "$($rtn.value) $($rtn.name) $($rtn.path)"
    }
}