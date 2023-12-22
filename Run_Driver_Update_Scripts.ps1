<#
.Synopsis
   Run Driver Update Checks for HP and Dell Workstations
.DESCRIPTION
   This script launches the driver update script for either Dell or HP depending on what's detected. It will exit if it's a different manufacturer than those
.EXAMPLE
   Example of how to use this cmdlet
.EXAMPLE
   Another example of how to use this cmdlet
#>

#--------------------MAIN BODY-----------------------------------------------------------
#----------------------------------------------------------------------------------------

Write-Host "Checking Computer Manufacturer"
$Manufactuer=(Get-CimInstance win32_ComputerSystem).Manufacturer

if ($Manufactuer -like "*Dell*") 
{
    Write-Host "Manufacturer is Dell. Launching Dell Driver Update Script"

    & "$PSScriptRoot\DellDriverUpdate.ps1"
    
}
elseif ($Manufactuer -like "*HP*") 
{
    Write-Host "Manufacturer is HP. Launching HP Driver Update Script"

    & "$PSScriptRoot\HPDriverUpdate.ps1"
}
else
{
    Write-Host "Manufacturer isn't supported, unable to run Asset Tag script." -ForegroundColor Red
    pause
}

remove-item -path "C:\Drivers" -Recurse
exit