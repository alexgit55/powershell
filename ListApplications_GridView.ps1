$32Bit = Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* -ErrorAction SilentlyContinue
$64Bit = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* -ErrorAction SilentlyContinue

$AppList = $32Bit + $64Bit

$FilteredList = @{}
foreach ($App in $AppList) {
    if ($null -eq $App.DisplayName) { continue }
    if ($FilteredList.ContainsKey($App.DisplayName)) {
        if ([version]$App.DisplayVersion -gt [version]$FilteredList[$App.DisplayName].DisplayVersion) {
            $FilteredList[$App.DisplayName] = $App
        }
    } else {
        $FilteredList.Add($App.DisplayName, $App)
    }
}

$FinalList = $FilteredList.Values | Select-Object DisplayName, DisplayVersion, UninstallString | Sort-Object DisplayName

$DisplayList = $FinalList | Out-GridView -PassThru

$DisplayList | Format-List
