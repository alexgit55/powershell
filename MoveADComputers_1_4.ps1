#Add-WindowsCapability –online –Name “Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0” | out-null

Import-Module ActiveDirectory

#--------------------Import Logging Functions and set the log path-----------------------
import-module write-log
#--------------------End Logging Information---------------------------------------------


#--------------------BEGIN FORM CONTROLS-------------------------------------------------

# Init PowerShell Gui
Add-Type -AssemblyName System.Windows.Forms
# Create a new form
$oulist=@()
$oulist = Get-ADOrganizationalUnit -LDAPFilter '(name=*)' -SearchBase 'ad path' -SearchScope Subtree | Sort-Object Name
 
$Form                    = New-Object system.Windows.Forms.Form
# Define the size, title and background color
$Form.ClientSize         = '900,750'
$Form.text               = "Alex Hill - Move Computers to OU"
$Form.BackColor          = [System.Drawing.ColorTranslator]::FromHtml("#44586a")
$Form.maximumsize = New-Object System.Drawing.Size(900,750)
$Form.MinimumSize = New-Object System.Drawing.Size(900,750)

#Text Label at the Top
$PCLabel = new-object System.Windows.Forms.Label
$PCLabel.Location = new-object System.Drawing.Size(10,20) 
$PCLabel.size = new-object System.Drawing.Size(100,20) 
$PCLabel.Text = "Enter The Computer Name(s) to Check or Move, one name per line:"
$PCLabel.Font =  New-Object System.Drawing.Font("Calibri",12,[System.Drawing.FontStyle]::Regular)
$PCLabel.AutoSize=$true

#TextBox for Computer Name
$PCName                     = New-Object system.Windows.Forms.TextBox
$PCName.multiline           = $true
$PCName.width               = 314
$PCName.height              = 100
$PCName.ScrollBars          = "Vertical"
$PCName.Visible             = $true
$PCName.location            = New-Object System.Drawing.Point(10,50)
$PCName.BackColor           = "#EEFBFB"
$PCName.Font =  New-Object System.Drawing.Font("Calibri",14,[System.Drawing.FontStyle]::Regular)

#Results Label
$RLabel = new-object System.Windows.Forms.Label
$RLabel.Location = new-object System.Drawing.Size(10,330) 
$RLabel.size = new-object System.Drawing.Size(100,20) 
$RLabel.Text = "Results:"
$RLabel.Font =  New-Object System.Drawing.Font("Calibri",12,[System.Drawing.FontStyle]::Regular)

#TextBox for Results
$Results                     = New-Object System.Windows.Forms.RichTextBox
$Results.multiline           = $true
$Results.width               = 800
$Results.height              = 250
$Results.ScrollBars          = "Vertical"
$Results.Visible             = $true
$Results.ReadOnly            = $true
$Results.location            = New-Object System.Drawing.Point(10,360)
$Results.BackColor           = "#EEFBFB"
$Results.Font =  New-Object System.Drawing.Font("Calibri",14,[System.Drawing.FontStyle]::Regular)

#Label for the OU DropDownList
$OULabel = new-object System.Windows.Forms.Label
$OULabel.Location = new-object System.Drawing.Size(10,160) 
$OULabel.size = new-object System.Drawing.Size(100,150) 
$OULabel.Text = "Select the Desired OU from the dropdown list (Only Used for Moving Computers):"
$OULabel.Font = "Calibri, 12"
$OULabel.AutoSize=$true

#Dropdown for OU List
$Combo = new-object System.Windows.Forms.ComboBox
$Combo.Location = new-object System.Drawing.Size(10,190)
$Combo.Size = new-object System.Drawing.Size(280,40)
$Combo.Width=400
$Combo.Font = New-Object System.Drawing.Font("Calibri",14,[System.Drawing.FontStyle]::Regular)
$Combo.BackColor           = "#EEFBFB"
#$Combo.AutoSize=$true

#Populate drop down list with OU's from active directory 
$ou_win10location=0
ForEach ($Item in $oulist) 
{
  $ou_win10location=$Item.DistinguishedName.IndexOf(",OU=Windows10")
  if ($ou_win10location -le 0) {continue}
  else
  {
    $ShortName=$Item.DistinguishedName.Substring(0,$ou_win10location)
    [void] $Combo.Items.Add($ShortName)
   }
}

#Label for Buttons
$ButtonLabel = new-object System.Windows.Forms.Label
$ButtonLabel.Location = new-object System.Drawing.Size(10,230) 
$ButtonLabel.size = new-object System.Drawing.Size(100,20) 
$ButtonLabel.Text = "Select Which Action You'd Like to Perform:"
$ButtonLabel.Font = "Calibri, 12"
$ButtonLabel.AutoSize=$true

# Inserting a Move button
$MoveButton = new-object System.Windows.Forms.Button
$MoveButton.Location = new-object System.Drawing.Size(170,260)
$MoveButton.Size = new-object System.Drawing.Size(150,60)
$MoveButton.Font = New-Object System.Drawing.Font('Calibri',12,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
$MoveButton.Text = "Move Computer(s)"
$MoveButton.Enabled=$false
$MoveButton.BackColor="#EEFBFB"
$MoveButton.Add_Click({MoveOU})

# Inserting a Cancel button
$CancelButton = new-object System.Windows.Forms.Button
$CancelButton.Location = new-object System.Drawing.Size(660,625)
$CancelButton.Size = new-object System.Drawing.Size(150,60)
$CancelButton.Font = New-Object System.Drawing.Font('Calibri',12,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
$CancelButton.Text = "Close"
$CancelButton.BackColor="#EEFBFB"
$CancelButton.DialogResult = [Windows.Forms.DialogResult]::OK
$Form.AcceptButton=$CancelButton

# Inserting a Check button
$CheckButton = new-object System.Windows.Forms.Button
$CheckButton.Location = new-object System.Drawing.Size(10,260)
$CheckButton.Size = new-object System.Drawing.Size(150,60)
$CheckButton.Font = New-Object System.Drawing.Font('Calibri',12,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
$CheckButton.Text = "Check Computer(s)"
$CheckButton.Enabled=$false
$CheckButton.BackColor="#EEFBFB"
$CheckButton.Add_Click({CheckOU})

$ClearButton                           = New-Object system.Windows.Forms.Button
$ClearButton.text                      = "Clear Output"
$ClearButton.width                     = 150
$ClearButton.height                    = 60
$ClearButton.location                  = New-Object System.Drawing.Point(660,285)
$ClearButton.Font                      = New-Object System.Drawing.Font('Calibri',12,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
$ClearButton.BackColor="#EEFBFB"
$ClearButton.Add_Click({Clear-Results})

$Form.controls.AddRange(@($MoveButton,$CancelButton,$Combo,$OULabel,$PCName,$PCLabel, $RLabel, $Results, $CheckButton, $ButtonLabel,$Clearbutton))


#--------------------END FORM CONTROLS-------------------------------------------------

#--------------------BEGIN FUNCTIONS---------------------------------------------------
#Checks for values inside textbox, disables select button if empty, enables once text is entered
function IsThereText
{
	if ($PCName.Text.Length -ne 0)
	{
		$MoveButton.Enabled = $true
        $CheckButton.Enabled=$true
	}
	else
	{
		$MoveButton.Enabled = $false
        $CheckButton.Enabled=$false
	}
}

function EndSession
{
    $Form.Close()
}

function CheckOU
{
    $Results.AppendText("Check Computer(s): ")
    $Results.AppendText("`r`n")
    #Pull the Computer Name from the Text Box
    ### Removing all the spaces and extra lines
    $x = $PCName.Lines | Where{$_} | ForEach{ $_.Trim() }
    ### Putting the array together
    $array = @()
    ### Putting each entry into array as individual objects
    $array = $x -split "`r`n"
    ForEach($computer in $array)
    {
      $FullPCDN=$null
      $FullPCDN=Get-ADComputer -Identity $computer -Properties DistinguishedName
      if ($FullPCDN -eq $null) 
      {
        $Results.SelectionColor = "#6D120E"
        $Results.AppendText("$Computer is not in Active Directory")
        $Results.AppendText("`r`n")
        $Results.AppendText("-------------------------------------------`r`n")
       }
      else
      {   
        $PCDN=$FullPCDN.ToString()
        $ResultLoc=$PCDN.IndexOf(",OU")
        $CurrentLocation=$PCDN.substring($ResultLoc+1)
        $Results.SelectionColor = "#6D120E"
        $Results.AppendText("$Computer is currently at $CurrentLocation")
        $Results.AppendText("`r`n")
        $Results.AppendText("-------------------------------------------`r`n")
       }
    }
    $Results.AppendText("`r`n")
    $Results.AppendText("Operation Completed")
    $Results.AppendText("`r`n")
    $Results.AppendText("*******************************************`r`n")
}

#Function to actually move computers to the desired OU
function MoveOU
{
    if ($Combo.SelectedItem -eq $null) 
    {
        $Results.AppendText("-------------------------------------------`r`n")
        $Results.AppendText("No OU selected, move will not proceed.")
        $Results.AppendText("`r`n")
        $Results.AppendText("-------------------------------------------`r`n")
        return
     }
    $Results.AppendText("Move Computer(s): ")
    $Results.AppendText("`r`n")
    $OUChoice=$Combo.SelectedItem.ToString()
    foreach ($ou in $oulist) 
    {
        $Win10loc=$ou.DistinguishedName.IndexOf(",OU=Windows10")
        if ($Win10loc -le 0) {continue}
        else
        {
            $Shorthand=$ou.DistinguishedName.Substring(0,$Win10loc)
            if ($Shorthand -eq $OUChoice) 
            {
                $DN=$ou.distinguishedname
                break
             }
        }
    }
    #Pull the Computer Name from the Text Box
    ### Removing all the spaces and extra lines
    $x = $PCName.Lines | Where{$_} | ForEach{ $_.Trim() }
    ### Putting the array together
    $array = @()
    ### Putting each entry into array as individual objects
    $array = $x -split "`r`n"

    ForEach($computer in $array)
    {
      $FullPCDN=$null
      $FullPCDN=Get-ADComputer -Identity $computer
      if ($FullPCDN -eq $null) 
      {
        $Results.SelectionColor = "#367355"
        $Results.AppendText("$Computer is not in Active Directory")
        $Results.AppendText("`r`n")
        $Results.AppendText("-------------------------------------------`r`n")
       }
      else 
      {
        Get-ADComputer $computer | Move-ADObject -TargetPath $DN
        $FullPCDN=Get-ADComputer -Identity $computer -Properties DistinguishedName
        $PCDN=$FullPCDN.ToString()
        $ResultLoc=$PCDN.IndexOf(",OU")
        $FinalLocation=$PCDN.substring($ResultLoc+1)
        $Results.SelectionColor = "#367355"
        $Results.AppendText("$Computer has been moved to $FinalLocation")
        $Results.AppendText("`r`n")
        $Results.AppendText("-------------------------------------------`r`n")
       }
    }
    $Results.AppendText("`r`n")
    $Results.AppendText("Operation Completed")
    $Results.AppendText("`r`n")
    $Results.AppendText("*******************************************`r`n") 
}

function Clear-Results {
    $Results.Text=""
    #$StatusLabel.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#FFC300")
    #$StatusLabel.Text = "Status: Output has been Cleared"
    Start-Sleep -Seconds 2

    #$StatusLabel.ForeColor= [System.Drawing.ColorTranslator]::FromHtml("#08E322")
    #$StatusLabel.Text = "Status: Ready"
}

$PCName.add_TextChanged({istheretext})

#--------------------END FUNCTIONS-------------------------------------------------

# ADD OTHER ELEMENTS ABOVE THIS LINE

# THIS SHOULD BE AT THE END OF YOUR SCRIPT FOR NOW
# Display the form
$Result = $Form.ShowDialog()

if ($result -eq [Windows.Forms.DialogResult]::OK) {
    write-host "Operation Complete"
    $Form.Dispose()
}
else {
    write-host "Operation cancelled"
    $Form.Dispose()
    exit
}

