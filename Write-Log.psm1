<#
.Synopsis
   Module to support logging during a script
.DESCRIPTION
   This module has functions that support logging when running a script.  It supports creating the log file/folder, writing to the log file 
   with several message types like error and information, and finally closing out the log
   It has a defaut log folder. If not logfoldername is supplied, it will save the log directly in the root
   It saves the log in csv format to be able to filter based on columns
#>

<#
.Synopsis
   Create the logfile and folder to start logging
.DESCRIPTION
   This Function creates the logfile and log folder if they don't aleady exist. It will append to an existing logfile.
   If no folder name is entered, then it will save the log directly in the root folder
.EXAMPLE
   Start-Log -LogName "Log Name" -FolderName "Log Folder"
.EXAMPLE
   Start-Log -LogName "Log Name" 
#>
function Start-Log
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)][string]$LogName,

        [Parameter()]
        [ValidateNotNullorEmpty()]
        [string]$FolderName="Blank"
    
    )

    #Default LogFolder path. Any name for the folder entered will be a subfolder under here
    $LogRoot="C:\Logs"

    #If no folder is enter and is left at the default of blank, it will set the logging folder to the default root. Otherwise the name
    #will be added to the root
    if ($FolderName -eq "Blank") 
    {
        $LogFolder=$LogRoot
    }
    else
    {
        $LogFolder="$LogRoot\$FolderName"
    }

    #This supports getting the current date the script was run at and adds the date to the name of the log in the formate of Name_Year_Month_Day
    $LogDate=get-date -Format FileDate
    $Log_Year=$LogDate.Substring(0,4)
    $Log_Month=$LogDate.Substring(4,2)
    $Log_Day=$LogDate.Substring(6,2)

    #If the log folder desired doesn't already exists, create the folder
    if (-not(test-path "$LogFolder")) {new-item -path "$LogFolder" -ItemType Directory | Out-Null}
    $script:LogPath="$LogFolder\$($LogName)_$($Log_Year)_$($Log_Month)_$($Log_Day).CSV"

    #Print where the logfile will be saved at to the screen
    Write-Host "Log File is located at $script:LogPath `n"

    #Standard logging format usage: Computer, Severity, Message. Since this is just beginning, fixed values will be used: Local machine, starting the log, message of starting script
    $ComputerName="Local"
    $Severity="Start"
    $Message="Beginning Script Execution: $LogName"

    Write-Host "Start: $Message `n"-ForegroundColor Green 

    #Custom object to write the logging format to the file: Current time, Computer Name, Severity Level and the Message
    [pscustomobject]@{
        Time = (Get-Date -f g)
        Computer = $ComputerName
        Severity = $Severity
        Message = $Message
    } | Export-Csv -Path $script:LogPath -Append -NoTypeInformation 
    
    #Optional sleep to slow down script execution so the messages can be read easier in real time
    Start-Sleep 1 
}

<#
.Synopsis
   Write log messages to the logfile created by Start-Log
.DESCRIPTION
   This functions writes a custom message to the logfile created by the start-log function as well as print the message to the console
   It supports various levels of severity including basic information, success, warning and error. it changes the write-host color
   depending on which severity is chosen. You can also enter in a computer name if it's a command targeting another device. If it's
   left blank, it will just write Local indicating it targeted the local machine
.EXAMPLE
   Write-Log "This is an information Message"
.EXAMPLE
   Write-Log "This is an error message" -Severity Error
.EXAMPLE
    Write-Log "This action targeted another computer" -ComputerName Computer
#>

function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]$Message,
 
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [ValidateSet('Information','Success','Warning','Error')]
        [string]$Severity = 'Information',

        [Parameter()]
        [ValidateNotNullorEmpty()]
        [string]$ComputerName='Local',

        [Parameter()]
        [ValidateSet('White','Gray')]
        [string]$InfoColor = 'White'

    )
    
    #Selects the text color based on the severity chosen for the Message
    switch ($Severity)
    {
        "Information" {Write-Host "$Message `n" -ForegroundColor "$InfoColor"}
        "Success" {Write-Host "Success: $Message `n" -ForegroundColor Cyan}
        "Warning" {Write-Host "Warning: $Message `n" -ForegroundColor Yellow}
        "Error" {Write-Host "Error: $Message `n" -ForegroundColor Red}
    }

    #Logging file format: Time, Computer, Severity, Message. Appends to Logfile.
    [pscustomobject]@{
        Time = (Get-Date -f g)
        Computer = $ComputerName
        Severity = $Severity
        Message = $Message
    } | Export-Csv -Path $script:LogPath -Append -NoTypeInformation

    Start-Sleep 1
 }


 <#
 .Synopsis
    Represents stopping the logging in the script
 .DESCRIPTION
    Writes specific text to the log file and console to represent that the script has finished
 #>
function Stop-Log
{   
    $ComputerName="Local"
    $Severity="Finish"
    $Message="Script Execution Finished"

    Write-Host "Finish: $Message `n"-ForegroundColor Green 

    [pscustomobject]@{
        Time = (Get-Date -f g)
        Computer = $ComputerName
        Severity = $Severity
        Message = $Message
    } | Export-Csv -Path $script:LogPath -Append -NoTypeInformation  

    Start-Sleep 1
}
