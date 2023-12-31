<#
.Synopsis
   Check for Security Agents on the Workstations
.DESCRIPTION
   The purpose of this script is to check for the required security agents and make sure they're installed on the workstation
#>

#--------------------Import Logging Functions and set the log path-----------------------
Import-Module "\\networkshare\modules\write-log"

Start-Log -LogName "Security Agent Check" -Folder "SecurityAgents"
#--------------------End Logging Information---------------------------------------------

#--------------------INITIAL VARIABLES---------------------------------------------------

$ServiceList=@('Agent1','Agent2','Agent3')

$global:NotInstalled=@()

$LaptopTypes=@(8, 9, 10, 11, 12, 14, 18, 21, 30, 31, 32)

$Technician=$env:USERNAME 

#--------------------FUNCTIONS-----------------------------------------------------------

#This Function checks to see that services in ServiceList are present on the workstation
function Check-Service
{
    Write-Log "Checking Service Status"
    
    $Services=(Get-Service).DisplayName
    foreach ($Item in $ServiceList)
    {
        if ($Services -match $Item) {Write-Log "$Item is installed" -Severity Success}
        else
        {
            Write-Log "$Item is not installed" -Severity Error
            $global:NotInstalled+=$Item
        }        
    }                    
}

#If any applications came up missing, then this function will run through the installers for each one
function Install-MissingApplications 
{
    Write-Log "Installing Missing Applications"
    foreach ($Application in $global:NotInstalled)
    {
        Write-Log "Now Installing $Application"
        switch ($Application)
        {
            'Agent1' {Install-Agent1}
            'Agent2' {Install-Agent2}
            'Agent3' {Install-Agent3}
        }
    }
}

#This Application runs the Installer for Agent1
function Install-Agent1
{
    $MSIArguments = @(
        "/i"
        "\\networkshare\msifiles.msi"
        "argument1"
        "arg2"
        "arg3"
        "/qn"
        "/norestart"
    )
    Start-Process "msiexec.exe" -ArgumentList $MSIArguments -Wait -NoNewWindow

    $CheckInstall=Get-Service "Agent1"

    if ($CheckInstall) {Write-Log "Agent1 has been installed successfully" -Severity Success }
    else {Write-Log "Agent1 was NOT installed successfully" -Severity Error}
}

#This Application runs the Installer for Agent2
function Install-Agent2
{
    $MSIArguments = @(
        "/i"
        "\\networkshare\msifiles.msi"
        "argument1"
        "arg2"
        "arg3"
        "/qn"
        "/norestart"
    )
    Start-Process "msiexec.exe" -ArgumentList $MSIArguments -Wait -NoNewWindow

    $CheckInstall=Get-Service "Agent2"

    if ($CheckInstall) {Write-Log "Agent2 has been installed successfully" -Severity Success }
    else {Write-Log "Agent2 was NOT installed successfully" -Severity Error}
}

#This Function runs through the installer for Agent3
function Install-Agent3
{
    $MSIArguments = @(
        "/i"
        "\\networkshare\msifiles.msi"
        "argument1"
        "arg2"
        "arg3"
        "/qn"
        "/norestart"
    )
    Start-Process "msiexec.exe" -ArgumentList $MSIArguments -Wait -NoNewWindow

    $CheckInstall=Get-Service "Agent3"

    if ($CheckInstall) {Write-Log "Agent1 has been installed successfully" -Severity Success }
    else {Write-Log "Agent3 was NOT installed successfully" -Severity Error}       
}

#--------------------MAIN BODY----------------------------
#---------------------------------------------------------

Check-Service

if ($global:NotInstalled -ne $null)
{
    Install-MissingApplications
}

Stop-Log