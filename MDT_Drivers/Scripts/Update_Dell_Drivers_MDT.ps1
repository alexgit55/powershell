<#
.Synopsis
   Script to update the Dell Drivers on an MDT Deployment Share
.DESCRIPTION
   This script takes in a deployment share path and model list. It will then download the latest drivers for those models from Dell and upload them into MDT.
   It will remove any drivers for that model currently there and upload the new ones downloaded
.EXAMPLE
   Example of how to use this cmdlet
.EXAMPLE
   Another example of how to use this cmdlet
#>

 [CmdletBinding()]
  Param(
     [Parameter(Mandatory)]
     [string]$DeploymentShare
 )

#Source Information and scripts from https://www.dell.com/support/kbdoc/en-us/000122176/driver-pack-catalog

#--------------------Import Logging Functions and set the log path-----------------------
Import-Module "C:\ModulePath\Write-Log.psm1"

Start-Log "Dell MDT Driver Update" -FolderName "MDT"
#--------------------End Logging Information---------------------------------------------

#--------------------Initial Variables---------------------------------------------------

$RootPath="C:\Drivers"

$CSVPath="$RootPath\Models"
$ModelList=Import-Csv -Path "$CSVPath\Dell_Driver_Models.csv"

$OS="Win10"
$Vendor="Dell Inc."

$VendorPath=$Vendor.Split(" ")
$VendorPath=$VendorPath[0]

$DriverRoot="$RootPath\$VendorPath"
if (!(Test-Path $DriverRoot)) {new-item -path $DriverRoot -ItemType Directory | Out-Null}

$ManifestRoot="\\NetworkShare\Manifests"
if (!(Test-Path -path "$ManifestRoot")) {New-Item -path "$ManifestRoot" -itemtype Directory | Out-Null}

$MDTModule = "C:\Program Files\Microsoft Deployment Toolkit\bin\MicrosoftDeploymentToolkit.psd1"
#----------------------END INITIAL VARIABLES--------------------------

#--------------------FUNCTIONS-----------------------------------------------------------

#Import MDT Module and Map Deployment Share
function Configure-MDT
{
    write-log -Message "Configuring MDT Settings"
    
    Import-Module $MDTModule -Scope Global
    
    if (!(Get-PSDrive -LiteralName PSDeploymentshare -ErrorAction 'silentlycontinue')) 
    {
	    Write-Log -Message "Adding $deploymentshare as a PSdrive"
        New-PSDrive -Name "PSDeploymentShare" -PSProvider MDTProvider -Root $deploymentshare -Scope Global | Out-Null
    }
    else
    {
	    Write-Log -Message "Removing PSDrive for Deployment Share"
        Remove-PSDrive PSDeploymentshare

        Write-Log -Message "Adding $deploymentshare as a PSdrive"
	    New-PSDrive -Name "PSDeploymentShare" -PSProvider MDTProvider -Root $deploymentshare -Scope Global | Out-Null
    }        
}


#Download Dell Driver Pack Catalog
function Download-DellDriverCab
{
    write-Log -Message "Downloading Latest Dell Driver Pack Cab"
    $DellDriverCabURL="http://downloads.dell.com/catalog/DriverPackCatalog.cab"

    $CabFolder="$DriverRoot\DriverPackCab"
    $CabDestination=$CabFolder + "\DriverPackCatalog.cab"

    if (test-path $CabDestination) {remove-item -Path $CabDestination -Force}

    if (!(Test-Path $CabFolder)) {new-item $CabFolder -ItemType Directory | Out-Null}
 
    Invoke-WebRequest -Uri $DellDriverCabURL -OutFile $CabDestination -ErrorAction SilentlyContinue | Out-Null       

    if (!(test-path $CabDestination)) {Write-Log -Message "Catlog Not able to be downloaded, exiting script" -Severity Error; exit}
    else {Write-Log -Message "Catalog downloaded to $CabFolder"}    
}

#Extracting XML File from CAB File
function Extract-DellDriverCab
{
    Write-Log -Message "Extracting XML catalog from Cab file"

    $CabFolder="$DriverRoot\DriverPackCab"

    $catalogCABFile = "$CabFolder" + "\DriverPackCatalog.cab"
    $catalogXMLFile = "$CabFolder" + "\DriverPackCatalog.xml"

    if (test-path -Path $catalogXMLFile) {remove-item -Path $catalogXMLFile -Force}

    EXPAND $catalogCABFile $catalogXMLFile | Out-Null
    if (test-path $catalogXMLFile) {Write-Log -Message "Catalog XML extracted to $CatalogXMLFile"}
    else {Write-Log -Message "Catalog XML unable to be extracted, exiting script" -Severity Error; exit}          
}

#Download Dell Drivers Based on the Selected Models
function Download-DellDrivers
{
    $CabFolder="$DriverRoot\DriverPackCab"
    $catalogXMLFile = "$CabFolder" + "\DriverPackCatalog.xml"
    
    write-log -Message "Loading XML File"
    [xml]$catalogXMLDoc = Get-Content $catalogXMLFile

    Write-Log -Message "Checking for Latest Driver Packs"
    foreach ($item in $ModelList)
    {
        $Model=$item.Model
        #When extracting drivers, it didn't work well with spaces in the name, so make a variable with the space removed. but still keep the full name in MDT for driver matching
        $ModelPath=$Model.Replace(" ","")

        Write-Log -Message "Current Model is $Model"
        $DriverPack=$catalogXMLDoc.DriverPackManifest.DriverPackage | ? { ($_.SupportedSystems.Brand.Model.name -eq "$Model") -and ($_.type -ne "WinPE") -and ($_.SupportedOperatingSystems.OperatingSystem.osArch -eq "x64") -and ($_.SupportedOperatingSystems.OperatingSystem.majorVersion -eq "10" ) -and ($_.SupportedOperatingSystems.OperatingSystem.minorVersion -eq "0" )}   
        
        #Check if a driver pack was found, if so, continue. Otherwise skip
        if ($DriverPack)
        {
            #If search returns multiple results, need to get the newest one only
            if ($DriverPack.Count -gt 1) 
            {
                $FinalPack=$DriverPack[0]
                foreach ($item in $DriverPack)
                {
                    $Date1=Get-Date -Date $FinalPack.dateTime
                    $Date2=Get-Date -Date $Item.dateTime
                    if ($Date2 -gt $Date1) {$FinalPack=$item}      
                }
                $DriverPack=$FinalPack
            }

            $PackDownloadLink = "http://" + $catalogXMLDoc.DriverPackManifest.baseLocation + $DriverPack.path
            $PackDownloadLink = "http://" + $catalogXMLDoc.DriverPackManifest.baseLocation + "/" + $DriverPack.path
            $Filename = [System.IO.Path]::GetFileName($PackDownloadLink)

            $DriverDownloadPath = "$DriverRoot\$ModelPath\Download"

            if (Test-Path $DriverDownloadPath) {remove-item $DriverDownloadPath -Recurse -Force}
            New-Item -path "$DriverRoot\$ModelPath" -Name "Download" -ItemType Directory | Out-Null

            $DriverPackDestination="$DriverDownloadPath\$FileName"
    
            Write-Log -Message "Downloading From $PackDownloadLink"

            Start-BitsTransfer -Source $PackDownloadLink -Destination $DriverPackDestination

            if (test-path $DriverPackDestination) {Write-Log -Message "Downloaded to $DriverPackDestination"; $global:DriverPackDownloads+=1}
             else {Write-Log -Message "Driver Pack not able to be downloaded" -Severity Error}
        }
        else
        {
            Write-Log -Message "No Driver Pack Available for $Model under current search conditions"
        }
    }
    Write-Log -Message "All Drivers Downloaded" -severity success       
}

#Extract Drivers from Downloaded Cab files
function Extract-DellDrivers
{
    Write-Log -Message "Beginning Extraction of Driver Files"
    foreach ($item in $ModelList)
    {
        $Model=$item.Model
        
        $ModelPath=$Model.Replace(" ","")
        $DriverDownloadPath = "$DriverRoot\$ModelPath\Download"
        
        if (!(test-path $DriverDownloadPath))
        {
            Write-Log -Message "Driver Pack for $Model was not downloaded, skipping to next Model" -Severity Warning
            continue
        }
        
        $DriverExtractPath="$DriverRoot\$ModelPath\Extract"
        if (!(test-path $DriverExtractPath)) {new-item $DriverExtractPath -ItemType Directory | Out-Null}
       
        write-log -Message "Extracting Drivers for $Model to $DriverExtractPath"
        $DriverFile=gci $DriverDownloadPath
        $Extension=$DriverFile.Extension

        #Run extraction commands based on whether driver pack is an exe or a cab file
        if ($Extension -eq ".exe") 
        {
            Write-Log -Message "Driver Pack is an exe file, running exe extraction"
            $Parms="/s /e=$DriverExtractPath"
            Start-Process $($DriverFile.FullName) -ArgumentList $Parms -Wait
            Remove-Item $($DriverFile.FullName) -Force
        }
        else
        {
            Write-Log -Message "Driver Pack is in a cab file, running cab extraction" 
            EXPAND -f:* $($DriverFile.FullName) $DriverExtractPath | Out-Null
            Remove-Item $($DriverFile.FullName) -Recurse
        }
    }
    write-Log -Message "Finished Driver Extraction"
        
}

#Extract the driver manifest for each model and save to separate location
function Extract-DriverManifest 
{
    Write-Log -Message "Starting Export of Driver Manifests"
    
    foreach ($item in $ModelList)
    {
        
        $Model=$item.Model
        
        $ModelPath=$Model.Replace(" ","")
        $DriverExtractPath="$DriverRoot\$ModelPath\Extract"
        
        if (!(test-path $DriverExtractPath)) {continue}
        
        Write-Log -Message "Exporting Manifest for $Model"
        
        $Manifest=gci -path "$DriverExtractPath" -Filter "manifest.xml" -Recurse
        $ManifestPath="$ManifestRoot\$Model"
        
        if ($Manifest)
        {
            Write-Log -Message "Saving Driver Manifest to $ManifestPath"
            if (!(test-path "$ManifestPath")) {new-item -Path "$ManifestPath" -ItemType Directory | out-null}
            copy-item -Path "$($Manifest.FullName)" -Destination "$ManifestPath" -Force
        }      
    }
}


#Import the Extracted Driver Files into MDT
function ImportDrivers-MDT
{
    write-Log -Message "Starting Driver Import into MDT"
    
    foreach ($item in $ModelList)
    {
        $Model=$item.Model
        
        $ModelPath=$Model.Replace(" ","")

        $DriverExtractPath="$DriverRoot\$ModelPath\Extract"
        if (!(test-path $DriverExtractPath)) {continue}
        
        write-log -Message "Delete and Recreate Folder for $Model"
        $oldpath = "PsDeploymentShare:\Out-of-Box Drivers\$OS\$Vendor\$Model"
		if (Test-Path $oldpath) { Remove-Item -Path $oldpath -Recurse -ErrorAction SilentlyContinue }
        if (!(Test-Path -path $oldpath)) {New-Item -Path "PsDeploymentShare:\Out-of-Box Drivers\$OS\$Vendor" -enable "True" -Name $Model -Comments "Dell Model" -ItemType Directory | Out-Null }

        Write-Log -Message "Importing Drivers for $Model"

        Import-MDTDriver -Path "PSDeploymentShare:\Out-of-Box Drivers\$OS\$Vendor\$Model" -SourcePath "$DriverExtractPath" -ImportDuplicates
        
        Write-Log -Message "Finished Importing for $Model" 
    }
    Write-Log -Message "Driver Import has Finished"
      
}

#Closing function of the script, empty folders and psdrive
function Script-Cleanup
{
    Write-Log -Message "Starting Script Cleanup"
    
    if (Get-PSDrive -LiteralName PSDeploymentshare -ErrorAction 'silentlycontinue') { Remove-PSDrive -Name "PSDeploymentShare" }
    gci -path $DriverRoot -Recurse | Remove-Item -recurse -force
    Write-Log -Message "Updated $DriverPackDownloads driver pack(s)"
    Stop-Log         
}

function Script-Startup
{
    Write-Log -Message "Beginning Script Execution"
    Write-Log -Message "Targeted Deployment share is $deploymentshare"
    Write-Log -Message "Driver Path is Located at $DriverRoot"   
}

#---------------------------------MAIN BODY--------------------------------------
Script-Startup

$global:DriverPackDownloads=0

Configure-MDT
Download-DellDriverCab
Extract-DellDriverCab
Download-DellDrivers

if ($global:DriverPackDownloads -gt 0)
{
    Write-Log -Message "At least one driver pack has been downloaded, running extraction and import"
    Extract-DellDrivers
    Extract-DriverManifest
    ImportDrivers-MDT
}
else
{
    Write-Log -Message "No Driver packs downloaded"
}

Script-Cleanup

