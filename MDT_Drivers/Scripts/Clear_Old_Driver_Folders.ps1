<#The purpose of this script is to clear out old driver folders in an MDT Deployment share
There is another script downloads and updates driver packs for models based on a list. It
will delete the current drivers for the models, and then import the latest ones. It will
be run on a monthly basis.  If the list changes and model is no longer getting updated, 
this script will detect that it hasn't been written to in a while and then delete them #>

 [CmdletBinding()]
  Param(
     [Parameter(Mandatory)]
     [string]$DeploymentShare
 )

#--------------------Import Logging Functions and set the log path-----------------------
Import-Module "\\NetworkShare\Write-Log.psm1"

Start-Log "MDT Driver CleanUp" -FolderName "MDT"
#--------------------End Logging Information---------------------------------------------

#--------------------Initial Variables---------------------------------------------------

$MDTModule = "C:\Program Files\Microsoft Deployment Toolkit\bin\MicrosoftDeploymentToolkit.psd1" #Location of the MDT Module

$OS="Win10" #The OS/foldername in MDT we want to work on

#The list of vendors we want to check the drivers out for
$VendorList=@("HP"
          "Dell Inc."
          "Microsoft Corporation")

#Will represent a running count of the driver sets removed to report on at the end
$global:DriverPackRemovals=0

#Current date that the script is run to use as reference for removal
$CurrentDate=Get-Date

#Date threshold to use for the check, anything older than this value in days will be removed
$AgeCheck="90"

#----------------------------------------------------------------------------------------

#--------------------FUNCTIONS-----------------------------------------------------------
#Import MDT Module and Map Deployment Share
function Configure-MDT
{
    write-log -Message "Configuring MDT Settings"
    
    Import-Module $MDTModule -Scope Global
    
    if (!(Get-PSDrive -LiteralName PSDeploymentshare -ErrorAction 'silentlycontinue')) 
    {
	    Write-Log -Message "Adding $deploymentshare as a PSdrive"
        New-PSDrive -Name "PSDeploymentShare" -PSProvider MDTProvider -Root $Deploymentshare -Scope Global | Out-Null
    }
    else
    {
	    Write-Log -Message "Removing PSDrive for Deployment Share"
        Remove-PSDrive PSDeploymentshare

        Write-Log -Message "Adding $deploymentshare as a PSdrive"
	    New-PSDrive -Name "PSDeploymentShare" -PSProvider MDTProvider -Root $Deploymentshare -Scope Global | Out-Null
    }        
}

#Remove HP Driver Folders that haven't been written to in 3 months
function Remove-OldDrivers
{
    
    Write-Log -Message "Clearing Old Driver Folders"

    foreach ($Vendor in $VendorList)
    {
        Write-Log -Message "Current Vendor is $Vendor"
        $DriverFolders=gci -path "PSDeploymentShare:\Out-of-Box Drivers\$OS\$Vendor"

        foreach ($Folder in $DriverFolders)
        {
            $FolderDate=get-date -Date $($Folder.CreatedTime)
            $TotalDays=$(New-TimeSpan -Start $FolderDate -End $CurrentDate).Days
        
            if ($TotalDays -gt $AgeCheck) 
            {
                try 
                {
                    Write-Log -Message "Drivers for $($Folder.Name) are $Totaldays days old and will be removed" -Severity Warning
                    $DeletePath="PsDeploymentShare:\Out-of-Box Drivers\$OS\$Vendor\$($Folder.Name)"
                    Remove-item -path $DeletePath -Recurse -ErrorAction Stop
                    Write-Log -Message "Drivers have been removed for $($Folder.Name)" -Severity Success
                    $global:DriverPackRemovals+=1
                }
                catch 
                {
                    Write-Log -Message "Drivers for $($Folder.Name) were not able to be deleted" -Severity Error
                }       
            }
            else {Write-Log -Message "Drivers for $($Folder.Name) are less than $AgeCheck days old and will not be removed"}
        }
    }
}

#Closing Function to remove psdrive, clean driver path and write closing items to log
function Script-Cleanup
{
    write-Log -Message "Starting Script Cleanup"
    
    if (Get-PSDrive -LiteralName PSDeploymentshare -ErrorAction 'silentlycontinue') { Remove-PSDrive -Name "PSDeploymentShare" }

    Write-Log -Message "Removed $global:DriverPackRemovals Model(s) today"
  
    Stop-Log         
}

#--------------------------MAIN BODY---------------------------------------

Configure-MDT

Remove-OldDrivers

Script-Cleanup

