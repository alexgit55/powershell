    $32Bit=@()
    $32Bit=Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*
    $64bit=@()
    $64bit=Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*
    $AppList=@()
    $AppList=$32Bit+$64bit

    $FinalList=@()
    foreach ($App in $AppList)
    {
        if ($App.DisplayName -eq $null) {continue}
        $FinalList=$FinalList += $App
    }

    $FinalList=$FinalList | select DisplayName, DisplayVersion, UninstallString | sort Displayname

    $DisplayList=$FinalList | Out-GridView -PassThru

    $DisplayList | Format-List
    