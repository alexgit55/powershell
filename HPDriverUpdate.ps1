<#
.Synopsis
   Run The HP Image Assistant (HPIA) to update drivers on a workstation
.DESCRIPTION
   This script will run the HPIA on the computer if it's not there already. This utility checks for the latest drivers
   for the workstation and installs them.  This only works on HP Workstations
.EXAMPLE
   Example of how to use this cmdlet
.EXAMPLE
   Another example of how to use this cmdlet
#>

#--------------------Import Logging Functions and set the log path-----------------------
Import-Module "\\modules\write-log"

Import-Module "\\modules\message-box"

Import-Module "\\Modules\Window-State"

Start-Log -LogName "HP Driver Update" -Folder "HPDriverUpdate"
#--------------------End Logging Information---------------------------------------------

#-------------------INITIAL VARIABLES----------------------------------------------------

$HPIA_Root="\HP Driver Updater\Files\HPImageAssistant"

$HPIA_Run="C:\HPDrivers\HPImageAssistant\HPImageAssistant.exe"

#--------------------FUNCTIONS-----------------------------------------------------------

#This function just checks whether it's an HP computer and continues if it is, otherwise it exits
function Check-Manufacturer
{
    $Manufactuer=(Get-CimInstance win32_ComputerSystem).Manufacturer
    
    if ($Manufactuer -like "*HP*") { Write-Log "Computer Manufacturer is HP, script can proceed"}
    else {Write-Log "Computer Manufacturer is Not HP, script cannot proceed" -Severity Error; pause; Stop-Log; exit}    
}

#This Function copies the program files locally in order to run them faster
function Copy-HPIAFiles 
{
    $HPIA_Copy="C:\HPDrivers"
    if (test-path "$HPIA_Copy") {Remove-Item -Path "$HPIA_Copy" -Recurse -Force}
    
    new-item -Path "$HPIA_Copy" -ItemType "Directory" | Out-Null
    
    Write-Log "Copying HP Image Assistant Files to the Computer"
    
    copy-item -path "$HPIA_Root" -Destination "$HPIA_Copy" -Recurse -Force

    Write-Log "Verifying HP Image Assistant copied to computer"

    if (test-path -Path "$HPIA_Run") {Write-Log "HPIA copied to the computer, script will proceed" -Severity Success}
    else {Write-Log "HPIA not copied to computer, script is not able to proceed." -Severity Error; pause; Stop-Log; exit}

}

#This function runs the HP Image assistant. It's set up to analyze the computer, download the relevant drivers and extract them. The extracted script will be then be run by another function
function Run-HPIA
{
    Write-Log "Beginning Driver Check and Install"

    $InstallCreated=$false

    Get-Process -Name *Powershell* | Set-WindowState -State MINIMIZE
    
    $HPIA_Args='/Operation:Analyze','/Category:All','/Selection:All','/Uwp:No','/Action:Extract','/noninteractive','/reportFolder:C:\OIT\HPDrivers','/softpaqdownloadfolder:C:\OIT\HPDrivers','/SoftpaqExtractFolder:C:\OIT\HPDrivers','/AutoCleanup'

    $hpia_result=Start $HPIA_Run -ArgumentList $HPIA_Args -Wait -PassThru

    Get-Process -Name *Powershell* | Set-WindowState -State RESTORE

    Write-Host "Return Code: $($HPIA_Result.ExitCode)"

    switch ($($HPIA_Result.ExitCode))
        {
            257 {$InstallCreated=$False}
            256 {$InstallCreated=$False}
            0   {if (test-path -Path "C:\OIT\HPDrivers\InstallAll.cmd"){$InstallCreated=$true}}
            Default {Write-Log "Error encountered in Running the application. Exit Code: $($HPIA_Result.ExitCode)" -Severity Error}
        } 
    return $InstallCreated  
}

#This functions runs the script created by the image assistant software in order to install all drivers downloaded and extracted
function Install-Drivers
{
    $DriversInstalled=$false 

    Write-Log "Beginning Driver Install"
    
    try 
    {
        Start cmd.exe -ArgumentList "/c C:\HPDrivers\InstallAll.cmd" -Wait -Verb RunAs
        Write-Log "Drivers Have been installed to the workstation" -Severity Success
        $DriversInstalled=$true
    }
    catch {write-log "Not able to run the install Script" -Severity Error}  

    return $DriversInstalled  
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
        remove-item -path "C:\HPDrivers" -Recurse
    }      
}


#--------------------END FUNCTIONS-------------------------------------------------------

#--------------------MAIN BODY-----------------------------------------------------------
$shell = New-Object -ComObject "Shell.Application"
$shell.minimizeall()

Get-Process -Name *Powershell* | Set-WindowState -State RESTORE

Check-Manufacturer

Copy-HPIAFiles
$InstallCreated=Run-HPIA

if ($InstallCreated -eq $true)
{
    Write-Log "The Driver Installation Script has been created" 
    $InstallDrivers=Install-Drivers
}
else
{
    Write-Log "The software did not find any recommendations for this workstation."
}

if ($InstallDrivers -eq $true)
{
    Write-Log "Driver installation has been completed, a restart is recommended to complete the installation" -Severity Warning
    Restart-Needed
}

Stop-Log

pause

$shell = New-Object -ComObject "Shell.Application"
$shell.UndoMinimizeALL()

remove-item -path "C:\HPDrivers" -Recurse