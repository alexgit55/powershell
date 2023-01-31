# Very little of this script is my work @nkofahl - Nathan Kofahl, HP inc. Most of it I stole with pride from people mentioned in these comments. 
# model tabel and download and extact logic from @gwblok - GARYTOWN.COM He built an excellent script for importing drivers into SCCM 
# HP Client Management Script Library https://ftp.hp.com/pub/caps-softpaq/cmit/hp-cmsl.html 

 [CmdletBinding()]
  Param(
     [Parameter(Mandatory)]
     [string]$DeploymentShare
 )

#--------------------Import Logging Functions and set the log path-----------------------
Import-Module "Write-Log.psm1"

Start-Log "HP MDT Driver Update" -FolderName "MDT"
#--------------------End Logging Information---------------------------------------------

#--------------------Initial Variables--------------------------------------------------- 
$OS = "Win10"
$OSVER = "21H2"
$Vendor="HP"

$RootPath="C:\Drivers"

$CSVPath="$RootPath\Models"
$ModelList=Import-Csv -Path "$CSVPath\HP_Driver_Models.csv"

$DriverRoot="$RootPath\$VendorPath"
if (!(Test-Path $DriverRoot)) {new-item -path $DriverRoot -ItemType Directory | Out-Null}

$ManifestRoot="\\NetworkShare\Manifests"
if (!(Test-Path -path "$ManifestRoot")) {New-Item -path "$ManifestRoot" -itemtype Directory | Out-Null}

$MDTModule = "C:\Program Files\Microsoft Deployment Toolkit\bin\MicrosoftDeploymentToolkit.psd1"

#Reset Vars
$DriverPack = ""
$Model = ""
#--------------------FUNCTIONS---------------------------------------------------

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

#Download Driver Packs
function Download-HPDrivers
{
    Write-Log -Message "Starting Driver Download"
    
    foreach ($item in $ModelList)
    {
        Write-Log -Message "Checking Model $($item.Model) Product Code $($item.ProdCode) for Driver Pack"
        
        $SavePath=$($item.Model).Replace(" ","")
	    $DriverPackDestination = "$DriverRoot\$SavePath"
        if (Test-Path $DriverPackDestination) {remove-item $DriverPackDestination -Recurse -Force}
        new-item $DriverPackDestination -ItemType Directory | Out-Null

        $DriverPack = New-HpDriverPack -platform $item.ProdCode -os $OS -osver $OSVER -Path $DriverPackDestination
       
       #Check if a Driver Was Found for the Current Model
        if ($DriverPack)
        {
            Write-Log -Message "Driver Pack for $($item.Model) downloaded to $DriverPackDestination"
            
            #Save the driver manifest file to a separate folder for reference
            $Manifest=gci -path "$DriverPackDestination" -Filter "manifest.xml" -Recurse
            $ManifestPath="$ManifestRoot\$($item.Model)"
            if ($Manifest)
            {
                Write-Log -Message "Saving Driver Manifest to $ManifestPath"
                if (!(test-path "$ManifestPath")) {new-item -Path "$ManifestPath" -ItemType Directory}
                copy-item -Path "$($Manifest.FullName)" -Destination "$ManifestPath" -Force
            }
            $global:DriverPackDownloads+=1
        }
        else
        {
            Write-Log -Message "No Driver Pack Available for $($item.Model) Product Code $($item.ProdCode) $($os) $($osver) via Internet" -Severity Warning
        }
    }
    write-Log -Message "All avaiable driver packs have been downloaded"
}


#Import Extracted Drivers into MDT
function ImportDrivers-MDT
{
    write-Log -Message "Starting Driver Import into MDT"
    foreach ($item in $ModelList)
    {
        $SavePath=$($item.Model).Replace(" ","")

        $DriverPackLocation="$DriverRoot\$SavePath"
        #Check to for Model Folder, if it doesn't exist, nothing was exported and can move on to next model
        if (!(test-path $DriverPackLocation)) {continue}

        write-log -Message "Delete and Recreate Folder for $Model"
        $oldpath = "PsDeploymentShare:\Out-of-Box Drivers\$OS\$Vendor\$($item.Model)"
		if (Test-Path $oldpath) { Remove-Item -Path "$oldpath" -Recurse -ErrorAction SilentlyContinue }
        New-Item -Path "PsDeploymentShare:\Out-of-Box Drivers\$OS\$Vendor" -enable "True" -Name "$($item.Model)" -Comments "HP Model" -ItemType "folder" | Out-Null
    
        Write-Log -Message "Importing Drivers for $($item.Model)"
        
        Import-MDTDriver -Path "PSDeploymentShare:\Out-of-Box Drivers\$OS\$Vendor\$($item.Model)" -SourcePath "$DriverPackLocation"    
    }
    
    Write-Log -Message "Drivers have been imported into MDT"  
}

#Closing Function to remove psdrive, clean driver path and write closing items to log
function Script-Cleanup
{
    write-Log -Message "Starting Script Cleanup"
    
    if (Get-PSDrive -LiteralName PSDeploymentshare -ErrorAction 'silentlycontinue') { Remove-PSDrive -Name "PSDeploymentShare" }
    gci -path $DriverRoot -Recurse | Remove-Item -recurse -force
    Write-Log -Message "Updated $global:DriverPackDownloads driver pack(s)"
    Stop-Log          
}

function Script-Startup
{
    Write-Log -Message "Targeted Deployment share is $deploymentshare"
    Write-Log -Message "Driver Path is Located at $DriverRoot"
}

#-----------------------------MAIN BODY---------------------------
#-----------------------------------------------------------------
Script-Startup

$global:DriverPackDownloads=0

Configure-MDT
Download-HPDrivers
#Check to see if a driver pack has been downloaded, if so, run extract and import, otherwise skip to end
if ($global:DriverPackDownloads -gt 0)
{
    Write-Log -Message "At least one driver pack has been downloaded, running import"
    ImportDrivers-MDT
}
else
{
    Write-Log -Message "No Driver packs needed updates"
}

Script-Cleanup
