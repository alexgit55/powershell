#------------------FUNCTIONS---------------------------

#Pause script execution until $Waittime. optional function if you want to have the script wait for a time before continuing
function Wait-ToStart ($WaitTime)
{
    #write-host "Start Sleeping until $WaitTime"
    Write-Host "Start Sleeping until $WaitTime"
    #(get-date "$WaitTime") - (get-date) | Start-Sleep

    (New-TimeSpan –End “$WaitTime”).TotalSeconds | Sleep
}


#List the deployment shares to update the drivers on, add or remove names as needed
function Initialize-ComboBox {
    #Populate drop down with list of commands desired 
    $CommandList=@(
                'DeploymentShare 1',
                'DeploymentShare 2'           
                )
    foreach ($item in $CommandList) {$Combo.Items.Add($item)}
    $Combo.SelectedIndex=0  
}

#The actual deployment shares that correspond to the names in the combobox functions. update as needed
function Get-DeploymentShare {
$ComboSelection=$Combo.SelectedItem.ToString()
switch ($ComboSelection)
    {
        "Archive Deployment Share" {$Choice="\\deploymentserver\deploymentshare1"}
        "Production Deployment Share" {$Choice="\\deploymentserver\deploymentshare2"}
    } 
    return $Choice   
}

<#---------------------FORM-------------------------------------------#>

Add-Type -assembly System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$monitor = [System.Windows.Forms.Screen]::PrimaryScreen

# Calculating the factors to multiply the Width and Height
$widthFactor  = 600 / $monitor.WorkingArea.Width
$heightFactor = 600 / $monitor.WorkingArea.Height

# MainForm

$MainForm = New-Object System.Windows.Forms.Form
$MainForm.text                       = "Update MDT Drivers - Production Server"
$MainForm.TopMost                    = $false
$MainForm.BackColor                  = [System.Drawing.ColorTranslator]::FromHtml("#44586a")

# using System.Drawing.Size will automatically round the width and height to integer numbers like with:
# [Math]::Round(1920 * $widthFactor)
# [Math]::Round(1080 * $heightFactor)
$MainForm.Size = New-Object System.Drawing.Size (1920 * $widthFactor), (1080 * $heightFactor)

$MainForm.StartPosition = "CenterScreen"
$MainForm.AutoSize = $true
$MainForm.BringToFront()
$MainForm.BackgroundImageLayout = "Stretch"
$MainForm.BackgroundImageLayout = "Stretch"

$InfoLabel                         = New-Object system.Windows.Forms.Label
$InfoLabel.text                    = "Ready to start? Select When, The Deployment Share and Click Begin"
$InfoLabel.AutoSize                = $true
#$InfoLabel.width                   = 100
$InfoLabel.height                  = 15
$InfoLabel.location                = New-Object System.Drawing.Point(15,15)
$InfoLabel.Font                    = New-Object System.Drawing.Font('Calibri',20)
$InfoLabel.ForeColor               = [System.Drawing.ColorTranslator]::FromHtml("#a2b4c3")
$MainForm.Controls.Add($InfoLabel)

# Create the collection of radio buttons
$RadioButton1                      = New-Object System.Windows.Forms.RadioButton
$RadioButton1.Location             = '20,90'
$RadioButton1.size                 = '360,30'
$RadioButton1.Checked              = $true 
$RadioButton1.Text                 = "Run With Delay until 9:00pm"
$RadioButton1.Font                    = New-Object System.Drawing.Font('Calibri',16)
$RadioButton1.ForeColor            = [System.Drawing.ColorTranslator]::FromHtml("#a2b4c3")
$RadioButton1.Add_Click({$CustomTime.Enabled=$false;$CustomTime.Visible=$false})
$MainForm.Controls.Add($RadioButton1)
 
$RadioButton2                         = New-Object System.Windows.Forms.RadioButton
$RadioButton2.Location                = '20,150'
$RadioButton2.size                    = '450,30'
$RadioButton2.Checked                 = $false
$RadioButton2.Text                    = "Run Driver Update Immediately"
$RadioButton2.Font                    = New-Object System.Drawing.Font('Calibri',16)
$RadioButton2.ForeColor               = [System.Drawing.ColorTranslator]::FromHtml("#a2b4c3")
$RadioButton2.Add_Click({$CustomTime.Enabled=$false;$CustomTime.Visible=$false})
$MainForm.Controls.Add($RadioButton2)

$RadioButton3                         = New-Object System.Windows.Forms.RadioButton
$RadioButton3.Location                = '20,210'
$RadioButton3.size                    = '550,30'
$RadioButton3.Checked                 = $false
$RadioButton3.Text                    = "Custom Start Time (ex 9:00pm OR 21:00)"
$RadioButton3.Font                    = New-Object System.Drawing.Font('Calibri',16)
$RadioButton3.ForeColor               = [System.Drawing.ColorTranslator]::FromHtml("#a2b4c3")
$RadioButton3.Add_Click({$CustomTime.Enabled=$True;$CustomTime.Visible=$true})
$MainForm.Controls.Add($RadioButton3)

$CustomTime                          = New-Object system.Windows.Forms.TextBox
$CustomTime.multiline                = $false
$CustomTime.Enabled                  = $false
$CustomTime.Visible                  = $false
$CustomTime.width                    = 100
$CustomTime.height                   = 70
$CustomTime.location                 = New-Object System.Drawing.Point(50,250)
$CustomTime.Font                     = New-Object System.Drawing.Font('Calibri',16)
$CustomTime.BackColor                = [System.Drawing.ColorTranslator]::FromHtml("#eefbfb")
$MainForm.Controls.Add($CustomTime)

$OkButton                          = New-Object system.Windows.Forms.Button
$OkButton.Enabled                  = $true
$OkButton.text                     = "Begin"
$OkButton.width                    = 150
$OkButton.height                   = 60
$OkButton.location                 = New-Object System.Drawing.Point(100,400)
$OkButton.Font                     = New-Object System.Drawing.Font('Calibri',14,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
$OkButton.BackColor                = [System.Drawing.ColorTranslator]::FromHtml("#9b9b9b")
$OkButton.DialogResult             = [Windows.Forms.DialogResult]::OK
$MainForm.AcceptButton=$OkButton
$MainForm.Controls.Add($OkButton)
    
$CancelButton                          = New-Object system.Windows.Forms.Button
$CancelButton.Enabled                  = $true
$CancelButton.text                     = "Cancel"
$CancelButton.width                    = 150
$CancelButton.height                   = 60
$CancelButton.location                 = New-Object System.Drawing.Point(300,400)
$CancelButton.Font                     = New-Object System.Drawing.Font('Calibri',14,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
$CancelButton.BackColor                = [System.Drawing.ColorTranslator]::FromHtml("#9b9b9b")
$CancelButton.DialogResult             = [Windows.Forms.DialogResult]::Cancel
$MainForm.CancelButton=$CancelButton
$MainForm.Controls.Add($CancelButton)

#Label for ComboBox
$ComboLabel                         = New-Object system.Windows.Forms.Label
$ComboLabel.text                    = "Select The Desired Deployment Share"
$ComboLabel.AutoSize                = $true
$ComboLabel.height                  = 15
$ComboLabel.location                = New-Object System.Drawing.Point(30,270)
$ComboLabel.Font                    = New-Object System.Drawing.Font('Calibri',16)
$ComboLabel.ForeColor               = [System.Drawing.ColorTranslator]::FromHtml("#a2b4c3")
$MainForm.Controls.Add($ComboLabel)

#Dropdown for OU List
$Combo = new-object System.Windows.Forms.ComboBox
$Combo.Location = new-object System.Drawing.Size(30,310)
$Combo.Size = new-object System.Drawing.Size(280,40)
$Combo.Width=250
$Combo.Font = New-Object System.Drawing.Font("Calibri",14,[System.Drawing.FontStyle]::Regular)
$Combo.BackColor = "#EEFBFB"
$MainForm.Controls.Add($Combo)

$MainForm.Add_Load({ Initialize-ComboBox })
 
#-----------------END FORM-----------------------------

#--------------------Main Body-----------------------------------------------------------

Write-Host "Beginning Script Execution - Update Dell and HP Drivers for MDT"

Set-Variable TimeDelay,WaitTime
$result = $MainForm.ShowDialog()

if ($result -eq [Windows.Forms.DialogResult]::OK) 
{
    if ($RadioButton1.Checked) {$TimeDelay=$true;$Waittime="9:00pm"}
    elseif ($RadioButton2.Checked) {$TimeDelay=$false}
    elseif ($RadioButton3.Checked) {$TimeDelay=$true;$Waittime=$CustomTime.Text}
    $DeploymentShare=Get-DeploymentShare
}
else {
    write-host ""
    write-host "Operation Cancelled, exiting script"
    write-host ""
    Start-Sleep 5
    exit
}

if ($TimeDelay -eq $true) 
{
    Wait-ToStart -WaitTime $WaitTime
    Write-Host "Sleep period passed" 
}

#Run Dell Driver Update Script

Write-Host "Launching Dell Driver Update Script"
    
& "$PSScriptRoot\Update_Dell_Drivers_MDT.ps1" -DeploymentShare "$DeploymentShare"
Write-Host "Dell Driver Update Has Finished" -Severity Success

start-sleep -Seconds 5

#Run HP Driver Update Script

Write-Host "Launching HP Driver Update Script"

& "$PSScriptRoot\Update_HP_Drivers_MDT.ps1" -DeploymentShare "$DeploymentShare"

Write-Host "HP Driver Update Has Finished" -ForegroundColor Cyan

start-sleep -Seconds 5

#Clear Non Updated Driver Folders

Write-Host "Launching Driver Cleanup Script"
    
& "$PSScriptRoot\Clear_Old_Driver_Folders.ps1"

Write-Host "Driver CleanUp Script Has Finished" -ForeGroundColor Cyan

start-sleep -Seconds 5
