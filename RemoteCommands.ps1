Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = New-Object System.Drawing.Point(775,600)
$Form.maximumsize = New-Object System.Drawing.Size(900,900)
$Form.MinimumSize = New-Object System.Drawing.Size(900,900)
$Form.text                       = "Alex - Run Commands on Remote Machines"
$Form.TopMost                    = $false
$Form.BackColor                  = [System.Drawing.ColorTranslator]::FromHtml("#44586a")

$PresetLabel                         = New-Object system.Windows.Forms.Label
$PresetLabel.text                    = "Select a Preset if Desired:"
$PresetLabel.AutoSize                = $true
#$PresetLabel.width                  = 250
$PresetLabel.height                  = 15
$PresetLabel.location                = New-Object System.Drawing.Point(15,15)
$PresetLabel.Font                    = New-Object System.Drawing.Font('Calibri',14)
$PresetLabel.ForeColor               = [System.Drawing.ColorTranslator]::FromHtml("#a2b4c3")

#Dropdown for Command List
$Combo = new-object System.Windows.Forms.ComboBox
$Combo.Location = new-object System.Drawing.Size(275,15)
$Combo.Size = new-object System.Drawing.Size(585,40)
$Combo.DropDownWidth=300
$Combo.Font = New-Object System.Drawing.Font("Calibri",14,[System.Drawing.FontStyle]::Regular)
$Combo.BackColor           = "#EEFBFB"
$Combo.DropDownStyle       ="DropDownList"
$Combo.AutoSize=$true
$combo.add_SelectedIndexChanged({Preset-ToCommandBox})

$CommandLabel                         = New-Object system.Windows.Forms.Label
$CommandLabel.text                    = "Type In Command to Run:"
$CommandLabel.AutoSize                = $true
#$PresetLabel.width                   = 250
$CommandLabel.height                  = 15
$CommandLabel.location                = New-Object System.Drawing.Point(15,60)
$CommandLabel.Font                    = New-Object System.Drawing.Font('Calibri',14)
$CommandLabel.ForeColor               = [System.Drawing.ColorTranslator]::FromHtml("#a2b4c3")

$CommandBox                        = New-Object system.Windows.Forms.TextBox
$CommandBox.Text                   = "C:\"
$CommandBox.multiline              = $true
$CommandBox.width                  = 585
$CommandBox.height                 = 100
$CommandBox.location               = New-Object System.Drawing.Point(275,60)
$CommandBox.Font                   = New-Object System.Drawing.Font('Calibri',14)
$CommandBox.BackColor              = [System.Drawing.ColorTranslator]::FromHtml("#eefbfb")

$PCLabel                         = New-Object system.Windows.Forms.Label
$PCLabel.text                    = "Enter Computer Name(s) to Target, one per line:"
$PCLabel.AutoSize                = $true
$PCLabel.width                   = 100
$PCLabel.height                  = 15
$PCLabel.location                = New-Object System.Drawing.Point(15,170)
$PCLabel.Font                    = New-Object System.Drawing.Font('Calibri',14)
$PCLabel.ForeColor               = [System.Drawing.ColorTranslator]::FromHtml("#a2b4c3")

$PCName                          = New-Object system.Windows.Forms.TextBox
$PCName.multiline                = $true
$PCName.width                    = 450
$PCName.height                   = 175
$PCName.location                 = New-Object System.Drawing.Point(15,200)
$PCName.Font                     = New-Object System.Drawing.Font('Calibri',14)
$PCName.BackColor                = [System.Drawing.ColorTranslator]::FromHtml("#eefbfb")
$PCName.add_TextChanged({IsThereText})

$RunButton                          = New-Object system.Windows.Forms.Button
$RunButton.Enabled                  = $false
$RunButton.text                     = "Run Command"
$RunButton.width                    = 150
$RunButton.height                   = 60
$RunButton.location                 = New-Object System.Drawing.Point(550,200)
$RunButton.Font                     = New-Object System.Drawing.Font('Calibri',14,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
$RunButton.BackColor                = [System.Drawing.ColorTranslator]::FromHtml("#9b9b9b")
$RunButton.Add_Click({Run-InVoke})

$CloseButton                           = New-Object system.Windows.Forms.Button
$CloseButton.text                      = "Close"
$CloseButton.width                     = 150
$CloseButton.height                    = 60
$CloseButton.location                  = New-Object System.Drawing.Point(710,200)
$CloseButton.Font                      = New-Object System.Drawing.Font('Calibri',14,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
$CloseButton.BackColor                 = [System.Drawing.ColorTranslator]::FromHtml("#9b9b9b")
$CloseButton.Add_Click({End-Session})

$SaveButton                           = New-Object system.Windows.Forms.Button
$SaveButton.text                      = "Save Output to File"
$SaveButton.width                     = 150
$SaveButton.height                    = 60
$SaveButton.location                  = New-Object System.Drawing.Point(710,350)
$SaveButton.Font                      = New-Object System.Drawing.Font('Calibri',14,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
$SaveButton.BackColor                 = [System.Drawing.ColorTranslator]::FromHtml("#9b9b9b")
$SaveButton.Add_Click({Save-Results})

$ClearButton                           = New-Object system.Windows.Forms.Button
$ClearButton.text                      = "Clear Output"
$ClearButton.width                     = 150
$ClearButton.height                    = 60
$ClearButton.location                  = New-Object System.Drawing.Point(550,350)
$ClearButton.Font                      = New-Object System.Drawing.Font('Calibri',14,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
$ClearButton.BackColor                 = [System.Drawing.ColorTranslator]::FromHtml("#9b9b9b")
$ClearButton.Add_Click({Clear-Results})

$Results                          = New-Object System.Windows.Forms.RichTextBox
$Results.multiline                = $true
$Results.width                    = 850
$Results.height                   = 350
$Results.location                 = New-Object System.Drawing.Point(15,475)
$Results.Font                     = New-Object System.Drawing.Font('Calibri',14)
$Results.ReadOnly                 = $true
$Results.BackColor                = "#D7DFD6"

$StatusLabel                         = New-Object system.Windows.Forms.Label
$StatusLabel.text                    = "Status: Ready"
$StatusLabel.AutoSize                = $true
$StatusLabel.width                   = 100
$StatusLabel.height                  = 15
$StatusLabel.location                = New-Object System.Drawing.Point(15,435)
$StatusLabel.Font                    = New-Object System.Drawing.Font('Calibri',14)
$StatusLabel.ForeColor               = [System.Drawing.ColorTranslator]::FromHtml("#08E322")

$Form.controls.AddRange(@($PresetLabel,$PCName,$PCLabel,$RunButton,$CloseButton,$Results,$Combo,$CommandBox,$CommandLabel,$StatusLabel,$SaveButton,$ClearButton))
$Form.Add_Load({ Initialize-ComboBox })

function Initialize-ComboBox {
    #Populate drop down with list of commands desired 
    $CommandList=@(
                'List Software Installs',
                'List Services and Status',
                'Restart Service (Change BITS to use a different service)',
                'List Java Paths',
                'Check Path Existence',
                'List Directory Items'
                'List .NET SDKS',
                'List .NET RunTimes'
                )
    foreach ($item in $CommandList) {$Combo.Items.Add($item)}
    $Combo.SelectedIndex=0  
}

function Convert-Combo {
$ComboSelection=$Combo.SelectedItem.ToString()
switch ($ComboSelection)
    {
        "List Software Installs" {$CommandString="${Function:Get-Applications}"}
        "List Services and Status" {$CommandString="Get-Service | Select Name, DisplayName, Status | sort DisplayName | format-table -AutoSize"}
        "Restart Service (Change BITS to use a different service)" {$CommandString="get-service -name BITS | restart-service"}
        "List Java Paths" {$CommandString="${Function:Get-JavaPaths}"}
        "Check Path Existence" {$CommandString="test-path -path '<Enter Path to Check>'"}
        "List Directory Items" {$CommandString="gci -path '<Enter Path to Check>' | select Name, Mode, LastWritetime | format-Table"}
        "List .NET SDKS" {$CommandString="dotnet --list-sdks"}
        "List .NET RunTimes" {$CommandString="dotnet --list-runtimes"}

    } 
    return $CommandString      
}

function Get-ComputerList {
    $Names = $PCName.Lines | Where{$_} | ForEach{ $_.Trim() }
    ### Putting the array together
    $Targets = @()
    ### Putting each entry into array as individual objects
    $Targets = $Names -split "`r`n"
    return $Targets
}

function Get-CommandBox {
    $ConvertCommand=$CommandBox.Text
  
    return $ConvertCommand
}

function Preset-ToCommandBox {
    $selected = Convert-Combo
    $CommandBox.Text = $selected
}

function Check-ComputerOnline($TargetComputer) {
    if (Test-Connection -ComputerName $TargetComputer -Quiet){ return $true }
    else {return $false}
}

function IsThereText {
	if ($PCName.Text.Length -ne 0)
	{
		$RunButton.Enabled = $true
	}
	else
	{
		$RunButton.Enabled = $false
	}
}

function Get-JavaPaths {
     $JavaSource='C:\Program Files\Java','C:\Program Files (x86)\Java'
     $JavaPaths=gci -path $JavaSource -ErrorAction SilentlyContinue
     $JavaFinal=@()
     foreach ($Path in $JavaPaths) {$JavaFinal+=$Path.FullName}
     return $JavaFinal
}



function Clear-Results {
    $Results.Text=""
    $StatusLabel.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#FFC300")
    $StatusLabel.Text = "Status: Output has been Cleared"
    Start-Sleep -Seconds 2

    $StatusLabel.ForeColor= [System.Drawing.ColorTranslator]::FromHtml("#08E322")
    $StatusLabel.Text = "Status: Ready"
}

function Save-Results {
    $ResultsData=$Results.Text

    if ($ResultsData -eq "")
    {
        $StatusLabel.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#FFC300")
        $StatusLabel.Text = "Status: Results currently Empty, no output saved"

        Start-Sleep -Seconds 5

        $StatusLabel.ForeColor= [System.Drawing.ColorTranslator]::FromHtml("#08E322")
        $StatusLabel.Text = "Status: Ready"
    }
    else 
    {
        $DateRun=get-date -Format FileDate
        if (-not (Test-path -Path "C:\RemoteResults")) {
            new-item -Path "C:\OIT" -Name RemoteResults -ItemType Directory
        }
        $ResultsData | out-file "C:\RemoteResults\RemoteResults_$DateRun.log" -Encoding utf8 -Append

        $StatusLabel.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#FFC300")
        $StatusLabel.Text = "Status: Output Saved to C:\RemoteResults\RemoteResults_$DateRun.log"
        Start-Sleep -Seconds 5

        $StatusLabel.ForeColor= [System.Drawing.ColorTranslator]::FromHtml("#08E322")
        $StatusLabel.Text = "Status: Ready"
    }

}

function Get-Applications {
    $ApplicationPaths='HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*','HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*'
    $AppList = Get-ItemProperty -Path $ApplicationPaths
    $FinalList=@()
    foreach ($App in $AppList)
    {
        if ($App.DisplayName -eq $null) {continue}
        $FinalList=$FinalList += $App
    }
   $FinalList=$FinalList | sort Displayname | format-list -Property DisplayName, DisplayVersion, InstallLocation, UninstallString
 
   return $FinalList
}

function Output-Error {
    if ($Error.Count -ge 1)
    {
        $Results.SelectionColor = "#9a2323"
        $Results.AppendText($Error)
        $Results.AppendText("`r`n")
        $Error.Clear()
    }

}

function Run-Invoke {
    $TargetList=Get-ComputerList
    $Command=Get-CommandBox
    $FullCommand=[Scriptblock]::Create("$Command") 

    $StatusLabel.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#E32308")
    $StatusLabel.Text = "Status: Processing..."
    
    $Results.AppendText("$(Get-Date) : Job Started")
    $Results.AppendText("`r`n")
    $Results.AppendText("Running: $Command")
    $Results.AppendText("`r`n")

    foreach ($Computer in $TargetList) {
        $Online=Check-ComputerOnline -TargetComputer $Computer
        if ($Online)
            {
                $Results.SelectionFont=New-Object System.Drawing.Font($Results.SelectionFont,'Bold')
                $Results.AppendText("$Computer : Online")
                $Results.AppendText("`r`n")
                $Results.AppendText("-------------------------------------------`r`n")
                try
                {
                    $Results.SelectionColor = "#6D120E"
                    $ResultsInfo=Invoke-Command -ComputerName $Computer -ScriptBlock $FullCommand | Out-String -Width 200
                    $Results.AppendText($ResultsInfo)
                    Output-Error
                    $Results.AppendText("`r`n")
                }
                catch
                {
                    $Results.AppendText("Command Failed")
                    $Results.AppendText("`r`n")
                    $Results.AppendText("-------------------------------------------`r`n")
                } 
            }
        else 
        {
                $Results.SelectionFont=New-Object System.Drawing.Font($Results.SelectionFont,'Bold')
                $Results.AppendText("$Computer : Offline")
                $Results.AppendText("`r`n")
                $Results.AppendText("-------------------------------------------`r`n")
                $Results.AppendText("`r`n")
        }
        
    }
    $StatusLabel.ForeColor= [System.Drawing.ColorTranslator]::FromHtml("#08E322")
    $StatusLabel.Text = "Status: Ready"
    $Results.AppendText("$(Get-Date) : Job Finished")
    $Results.AppendText("`r`n")
    $Results.AppendText("--------------------------------------------------------------------------------------`r`n")
    $Results.AppendText("`r`n")
}

function End-Session {
    $Form.Close()
}

[void]$Form.ShowDialog()

$Form.Dispose()
