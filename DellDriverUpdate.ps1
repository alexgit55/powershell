<#
.Synopsis
   Run The Dell Command Update Utility (DCU) to update drivers on a workstation
.DESCRIPTION
   This script will install the DCU on the computer if it's not there already. This utility checks for the latest drivers
   for the workstation and installs them.  This only works on Dell workstations
.EXAMPLE
   Example of how to use this cmdlet
.EXAMPLE
   Another example of how to use this cmdlet
#>

#--------------------Import Logging Functions and set the log path-----------------------
Import-Module "\modules\write-log"

Import-Module "\modules\message-box"

import-module "\Modules\Window-State"

Start-Log -LogName "Dell Driver Update" -Folder "DellDriverUpdate"
#--------------------End Logging Information---------------------------------------------

#-------------------INITIAL VARIABLES----------------------------------------------------

$CheckVersion="4.7.1"

$DCUInstaller="\\\Dell Driver Updater\Files\Installer\DellCommandUpdate_4_7.msi"

$DCUPath="C:\Program Files\Dell\CommandUpdate\dcu-cli.exe"

#--------------------FUNCTIONS-----------------------------------------------------------

#This function just checks whether it's a Dell computer and continues if it is, otherwise it exits
function Check-Manufacturer
{
    $Manufactuer=(Get-CimInstance win32_ComputerSystem).Manufacturer
    
    if ($Manufactuer -like "*Dell Inc.*") { Write-Log "Computer Manufacturer is Dell, script can proceed"}
    else {Write-Log "Computer Manufacturer is Not Dell, script cannot proceed" -Severity Error; pause; Stop-Log; exit}    
}

#This function checks to see if DCU is installed already
function Check-ForDCU
{
    Write-log "Checking for Dell Command Update Software"
    
    $DCUCheck=Get-WmiObject -Class Win32_Product | Where-Object{$_.Name -like "*Dell Command | Update*"}
    $DCUVersion=$DCUCheck.Version
    return $DCUVersion        
}

#This Function removes Dell Command Update from the computer
function Remove-DCU
{
    Write-Log "Uninstalling Dell Command Update software"
    
    $DCUCheck=Get-WmiObject -Class Win32_Product | Where-Object{$_.Name -like "*Dell Command | Update*"}
    $DCUCheck.Uninstall() | Out-Null
          
}


#This function installs the Dell Command Update software
#It will call the function to check if installed already first, then call the function to remove if an older version
function Install-DCU 
{   
    $DCUVersion=Check-ForDCU
    if ($DCUVersion)
    {
         if ($DCUVersion -eq $CheckVersion)
         {
            Write-Log "Updated Version of Dell Command Update already present, no need to install" -Severity Success
            return
         }
         else
         {
            Write-Log "Different Version of Dell Command Update present, will remove and install current version" -Severity Warning
            try
            {
                Remove-DCU
                $DCUVersion=Check-ForDCU
                if ($DCUVersion -eq $null) {Write-Log "Dell Command Update Has Been UnInstalled"}
            }
            catch {Write-Log "Dell Command Update Not able to be uninstalled" -Severity Error}
         }
     }
     else {Write-Log "Dell Command Update Not Installed, will proceed with Installation"}

     Write-Log "Installing Dell Command Update"

     start-process msiexec.exe -ArgumentList "/i `"$($DCUInstaller)`" /qb" -Wait

     $DCUVersion=Check-ForDCU
     if ($DCUVersion) { Write-Log "Dell Command Update $DCUVersion has been installed successfully"}
     else {Write-Log "Dell Command Update not able to be installed. Script cannot proceed"; exit}
                      
}

#This function sets some initial configuration options for the Dell Command Update software
#The settings are setting the driver download location, to auto suspend bitlocker if required, and prevent it from running automatically

function Configure-DCU
{
    Write-Log "Beginning DCU Configuration"
    
    $Args1='/configure','-downloadlocation=C:\OIT\DellDrivers'
    $Args2='/configure','-autoSuspendBitLocker=enable'
    $Args3='/configure','-scheduleManual'

    try
    {
        & $DCUPath $Args1
        & $DCUPath $Args2
        & $DCUPath $Args3
        
        Write-Host " "
    }
    catch
    {
        Write-Log "Error Configuring application. Unable to proceed" -Severity Error
    }

}

#This Function runs the DCU software to check for and install drivers for the current workstation
function Run-DCU
{
    Write-Host " "

    Write-Log "Beginning Driver Check and Install"

    $RestartNeeded=$false

    try
    {
        Get-Process -Name *Powershell* | Set-WindowState -State MINIMIZE
        
        $DCU_Args='/applyUpdates','-updateType=driver,bios,firmware','-reboot=Disable','-outputLog=C:\OIT\Logs\DellDriverUpdate\DCU_4_7.log'

        $DCU_Result=Start $DCUPath -ArgumentList $DCU_Args -Wait -PassThru

        Get-Process -Name *Powershell* | Set-WindowState -State RESTORE

        Write-Log "Return Code: $($DCU_Result.ExitCode)"

        switch ($($DCU_Result.ExitCode))
        {
            500 {$RestartNeeded=$false}
            5   {$RestartNeeded=$True}
            1   {$RestartNeeded=$True}
            0   {$RestartNeeded=$false}
            Default {Write-Log "Error encountered in Running the application. Exit Code: $($DCU_Result.ExitCode)" -Severity Error}
        }   

    }
    catch
    {
        Write-Log "Error encountered in Running the application.  Exit Code: $($DCU_Result.ExitCode)" -Severity Error
    }

   return $RestartNeeded
}

#This Function will prompt to restart the computer
function Restart-Needed
{
    $Restart=Show-MessageBox -Title "Restart Computer" -Message "A Restart has been recommended. Press OK to restart or Cancel to stop. This will Timeout in 30 seconds and default to restarting the computer" -Buttons OKCancel -Timeout 30 -MinimizeWindows -icon Exclamation
    
    if ($Restart -eq "Cancel")
    {
        Write-Log "Restart Will NOT occur. Please restart soon to apply updates and Suspend Bitlocker if needed" -Severity Warning
        start-sleep 2   
    }
    else
    {
        Write-Log "Restart of the computer has been requested" -Severity Warning
        Write-Log "Suspending bitlocker for BIOS if needed"       
        Suspend-BitLocker -MountPoint "C:" 
        Write-host " "
        Write-Log "Computer will Be Restarted in 10 seconds" -Severity Warning
        Start-Sleep 10
        Stop-Log
        Restart-Computer -Force
        remove-item -path "C:\OIT\HPDrivers" -Recurse
    }      
}

#--------------------END FUNCTIONS-------------------------------------------------------

#--------------------MAIN BODY-----------------------------------------------------------
$shell = New-Object -ComObject "Shell.Application"
$shell.minimizeall()

Get-Process -Name *Powershell* | Set-WindowState -State RESTORE

Check-Manufacturer
Install-DCU
Configure-DCU
$RestartNeeded=Run-DCU

if ($RestartNeeded -eq $true)
{
    Write-Log "The Application is recommending a restart to complete" -Severity Warning
    Remove-DCU
    Restart-Needed
}
else
{
    Write-Log "No Restart is needed"
    Remove-DCU
}

Stop-Log

pause

$shell = New-Object -ComObject "Shell.Application"
$shell.UndoMinimizeALL()

remove-item -path "C:\OIT\DellDrivers" -Recurse