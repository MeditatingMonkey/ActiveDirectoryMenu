param (
    [int]$SelectedIndex,
    [bool]$AddPrimaryDevice,
    [bool]$UseUsernameForLog
)

# Load the System.Web assembly
Add-Type -AssemblyName "System.Web"
Add-Type -AssemblyName System.Windows.Forms

Function Get-UserComp ($username) {
    Write-Host "Username : $username"
    
    $SCCMPSModule = "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1"
    Import-Module ActiveDirectory
    Import-Module -Name $SCCMPSModule

    $SiteCode = 'FEI'
    Set-Location "$($SiteCode):"

    $userDeviceaffinity = Get-CMUserDeviceAffinity -UserName $username | Select-Object ResourceName
    
    if ($userDeviceaffinity -eq $null) {
        Write-Host "No Primary Device Found! Switch to another OU from the Droupdown list."
    } else {
        return $userDeviceaffinity.ResourceName
    }
}

Function Draw_Main_Form
{
    #Draw the main window
    $Form = New-Object System.Windows.Forms.Form
    $Form.Text = "Software Deployment"
    $Form.Size = New-Object System.Drawing.Size(800, 525)
    $Form.StartPosition = "CenterScreen"

    #SELECT OU LABEL
    $OULabel = New-Object System.Windows.Forms.Label
    $OULabel.Location = New-Object System.Drawing.Point(25, 15) 
    $OULabel.Text = "Select OU:"
    $OULabel.Font = New-Object System.Drawing.Font('Verdana Pro',10)
    $OULabel.AutoSize = $true

    #DROP DOWN LIST
    $ComboBoxOU = New-Object System.Windows.Forms.ComboBox
    $ComboBoxOU.Location = New-Object System.Drawing.Point(25, 35)
    $ComboBoxOU.Size = New-Object System.Drawing.Size(350, 20)
    $ComboBoxOU.Font = New-Object System.Drawing.Font('Verdana Pro',10)
    $ComboBoxOU.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList

    #To Show the last OU name in the dropdown list
    foreach ($OU in $OUs) {   
        $LastOUName = [regex]::Match($OU, '(?<=OU=)[^,]+').Value  
        $ComboBoxOU.Items.Add($LastOUName)
    }
    
    if ($AddPrimaryDevice) {
        $ComboBoxOU.Items.Add("Primary Device")
    }

    $ComboBoxOU.SelectedIndex = $SelectedIndex

    #Search Computer Name: Label
    $SearchCompLabel = New-Object System.Windows.Forms.Label
    $SearchCompLabel.Location = New-Object System.Drawing.Point(25, 70) 
    $SearchCompLabel.Text = "Search Computer Name: "
    $SearchCompLabel.Font = New-Object System.Drawing.Font('Verdana Pro',10)
    $SearchCompLabel.AutoSize = $true

    #Search Box - Computer
    $SearchBoxComp = New-Object System.Windows.Forms.TextBox
    $SearchBoxComp.Size = New-Object System.Drawing.Size(350, 20)
    $SearchBoxComp.Location = New-Object System.Drawing.Point(25, 90)
    $SearchBoxComp.Font = New-Object System.Drawing.Font('Verdana Pro',10)

    #List Box Show the list of devices
    $ListBoxComp = New-Object System.Windows.Forms.ListBox
    $ListBoxComp.Size = New-Object System.Drawing.Size(350, 300)
    $ListBoxComp.Location = New-Object System.Drawing.Point(25, 115)
    $ListBoxComp.Font = New-Object System.Drawing.Font('Verdana Pro',10)
    $ListBoxComp.SelectionMode = 'Multiextended'   #Enables multi selection
    $ListBoxComp.Items.Clear()
    $script:CompNames = @()  #The CompNames Array is being initialised and $script means that it is accessible from variables outside the fucntions and can be called from anywhere in the script.
    
    function UpdateCompNames {
        $username = "CORP\" + $username    
        if ($ComboBoxOU.SelectedItem -eq "Primary Device") {
            $script:CompName = Get-UserComp $username
            $script:CompNames = @($script.CompName)
        } else {
        $script:OUCOMP = $OUs[$ComboBoxOU.SelectedIndex] #Check the Option selected from the DDL
        $script:CompNames = Get-ADComputer -LDAPFilter "(name=FBC*)" -SearchBase $OUCOMP | Sort-Object Name #Filter the computers with FBC name
        }
        UpdateListBox 
    }

    $ComboBoxOU.Add_SelectedIndexChanged({
        UpdateCompNames
    })

    function UpdateListBox {
        $SearchTextComp = $SearchBoxComp.Text #Allocating the Searched text to the variable
        $WildcardSearchTextComp = "*$SearchTextComp*"

        if ($ComboBoxOU.SelectedItem -eq "Primary Device") {
            if ($script:CompName -like $WildcardSearchTextComp) {
                $FilteredComps = @($script:CompName)
            } else {
                $FilteredComps = @()
            }
        } else {
            $FilteredComps = $CompNames | Where-Object { ($_.Name -like $WildcardSearchTextComp) } #String approximation
        }
        
        $ListBoxComp.Items.Clear() #refreshing the list

        foreach ($FilteredComp in $FilteredComps) {
            if ($ComboBoxOU.SelectedItem -eq "Primary Device") {
                $ListBoxComp.Items.Add($FilteredComp) 
            } else {
                $ListBoxComp.Items.Add($FilteredComp.Name)
            }
        }
    }

    $SearchBoxComp.Add_TextChanged({
        UpdateListBox
    })

    # Update CompNames for the initial population
    UpdateCompNames

    $ADGroups = Get-ADGroup -Filter * -SearchBase $OUApp | Sort-Object Name #Load the Variable with AD Groups

    #Search Application Name Label
    $SearchAppLabel = New-Object System.Windows.Forms.Label
    $SearchAppLabel.Location = New-Object System.Drawing.Point(385, 70) 
    $SearchAppLabel.Text = "Search Application Name: "
    $SearchAppLabel.Font = New-Object System.Drawing.Font('Verdana Pro',10)
    $SearchAppLabel.AutoSize = $true

    #Search Box - Applications
    $SearchBox = New-Object System.Windows.Forms.TextBox
    $SearchBox.Size = New-Object System.Drawing.Size(375, 20)
    $SearchBox.Location = New-Object System.Drawing.Point(385, 90)
    $SearchBox.Font = New-Object System.Drawing.Font('Verdana Pro',10)

    #List box to Show the list of applications
    $ListBox = New-Object System.Windows.Forms.ListBox
    $ListBox.Size = New-Object System.Drawing.Size(375, 145)
    $ListBox.Location = New-Object System.Drawing.Point(385, 115)
    $ListBox.Font = New-Object System.Drawing.Font('Verdana Pro',10)

    # Add the AD groups to the ListBox - Hide Available and "AP_SCCM_FBC_"
    foreach ($ADGroup in $ADGroups) 
    {
        if (-not $ADGroup.Name.EndsWith("Available"))
        {
            $ListBox.Items.Add($ADGroup.Name.Replace("AP_SCCM_FBC_", "")) #splitting the names with prefix
        }
    }

    $SearchBox.Add_TextChanged({
        $SearchText = $SearchBox.Text #Allocating the Searched text to the variable
        $WildcardSearchText = "*$SearchText*"
        $FilteredGroups = $ADGroups | Where-Object { ($_.Name -like $WildcardSearchText) -and (-not $_.Name.EndsWith("Available")) } # - Hide Available AD groups
        $ListBox.Items.Clear()

        foreach ($FilteredGroup in $FilteredGroups) 
        {
            $ListBox.Items.Add($FilteredGroup.Name.Replace("AP_SCCM_FBC_", ""))
        }
    })

    $SelectedAppLabel = New-Object System.Windows.Forms.Label
    $SelectedAppLabel.Location = New-Object System.Drawing.Point(383, 253) 
    $SelectedAppLabel.Text = "Selected Application Name: "
    $SelectedAppLabel.Font = New-Object System.Drawing.Font('Verdana Pro',10)
    $SelectedAppLabel.Size = New-Object System.Drawing.Size(174, 20)

    # Button to Clear Selected Item from $SelectedListBox 
    $ClearSelectedItemButton = New-Object System.Windows.Forms.Button
    $ClearSelectedItemButton.Location = New-Object System.Drawing.Point(560, 250)
    $ClearSelectedItemButton.Size = New-Object System.Drawing.Size(120, 20)
    $ClearSelectedItemButton.Text = "Clear Selected"
    $ClearSelectedItemButton.Font = New-Object System.Drawing.Font('Verdana Pro', 10)

    $ClearSelectedItemButton.Add_Click({
        if ($SelectedListBox.SelectedItems.Count -gt 0) {
            $SelectedItems = $SelectedListBox.SelectedItems | ForEach-Object { $_ }

            foreach ($item in $SelectedItems) {
                $SelectedListBox.Items.Remove($item)
            }
        } else {
            [System.Windows.Forms.MessageBox]::Show("Please select at least one item to remove.", "Warning", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
    })

    # Button to Clear All Items from $SelectedListBox
    $ClearAllItemsButton = New-Object System.Windows.Forms.Button
    $ClearAllItemsButton.Location = New-Object System.Drawing.Point(690, 250)
    $ClearAllItemsButton.Size = New-Object System.Drawing.Size(70, 20)
    $ClearAllItemsButton.Text = "Clear All"
    $ClearAllItemsButton.Font = New-Object System.Drawing.Font('Verdana Pro', 10)

    $ClearAllItemsButton.Add_Click({
        $SelectedListBox.Items.Clear()
    })

    $SelectedListBox = New-Object System.Windows.Forms.ListBox
    $SelectedListBox.Size = New-Object System.Drawing.Size(375, 145)
    $SelectedListBox.Location = New-Object System.Drawing.Point(385, 274)
    $SelectedListBox.Font = New-Object System.Drawing.Font('Verdana Pro',10)
    $SelectedListBox.SelectionMode = "Multiextended"

    $ListBox.Add_SelectedIndexChanged({
        $SelectedItem = $ListBox.SelectedItem

        if ($SelectedItem -and -not ($SelectedListBox.Items.Contains($SelectedItem))) {
            $SelectedListBox.Items.Add($SelectedItem)
        }
    })

    ### Adding an OK button to the text box window
    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Point(685,420) ### Location of where the button will be
    $OKButton.Size = New-Object System.Drawing.Size(75,23) ### Size of the button
    $OKButton.Text = 'OK' ### Text inside the button
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $Form.AcceptButton = $OKButton
    $Form.Controls.Add($OKButton)
                    
    #Cancel Button
    ### Adding a Cancel button to the text box window
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Point(25,420) ### Location of where the button will be
    $CancelButton.Size = New-Object System.Drawing.Size(75,23) ### Size of the button
    $CancelButton.Text = 'Cancel' ### Text inside the button
    $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $Form.CancelButton = $CancelButton
    $Form.Controls.Add($CancelButton)

    # ... Code to create the form and add controls ...
    $Form.Controls.AddRange(@($OUlabel,$SearchCompLabel,$ComboBoxOU,$SearchBoxComp,$ListBoxComp,$SearchAppLabel,$SearchBox,$ListBox,$SelectedAppLabel, $SelectedListBox, $ClearSelectedItemButton, $ClearAllItemsButton, $OKButton, $CancelButton))

    #Show the Main Form
    $result = $Form.ShowDialog()
    Form_Results
}


Function Form_Results {
    # Process the results of the main form
    # If Retry selected

    if ($result -eq [System.Windows.Forms.DialogResult]::Retry)
    {
        Search
    }

    # If Ok selected
    if ($result -eq [System.Windows.Forms.DialogResult]::OK)
    {
        # Get The Computer Name
        $ComputerToAdd = $ListBoxComp.SelectedItems
        # Get the AD Group Names
        $GroupsToAdd = $SelectedListBox.Items

        # Check to see if the computers are empty. If empty exit program so they can do better next time
        If ($ComputerToAdd.count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Please select at least one computer to continue.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            Write-Host "No computers selected. Retrying."
            Draw_Main_Form
        }

        # Check to see if the ad groups are empty. If empty exit program so they can do better next time
        elseif ($GroupsToAdd.count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Please select at least one AD group to continue.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            Write-Host "No AD Groups selected. Retrying"
            Draw_Main_Form
        }
    }

    # If the cancel selected
    if ($result -eq [System.Windows.Forms.DialogResult]::Cancel)
    {
        Write-Host "Exiting Software Deployer."
        Exit
    }
    
    # Display a confirmation prompt with the selected computer names and AD groups
    $message = "You are about to add the following Computer(s) to the selected AD Group(s):`n`n"
    $message += "Computers:`n"  

    foreach ($computer in $ComputerToAdd) {
        $message += "- $computer`n" #iteration
    }

    $message += "`nGroups:`n"

    foreach ($group in $GroupsToAdd) {
        $message += "- $group`n" #iteration
    }

    $confirmation = [System.Windows.Forms.MessageBox]::Show($message, "Confirmation", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    
    if ($confirmation -eq [System.Windows.Forms.DialogResult]::Yes)
    {
        $prefix = "AP_SCCM_FBC_" #Adding back the prefix

        $GroupsAdded = @()

        foreach ($group in $GroupsToAdd) {
            # Check if group name starts with DM, SCCM or RL
            if ($group.StartsWith("DM") -or $group.StartsWith("SCCM") -or $group.StartsWith("RL")) {
                $groupWithPrefix = $group
            } else {
                $groupWithPrefix = $prefix + $group
            }
            foreach ($computer in $ComputerToAdd) {
                $adComputer = Get-ADComputer -Identity $computer

                $groupMembership = Get-ADPrincipalGroupMembership -Identity $adComputer | Where-Object { $_.DistinguishedName -eq $adGroup.DistinguishedName }

                if ($groupMembership) {
                    [System.Windows.Forms.MessageBox]::Show("The computer $computer is already a member of the group $groupWithPrefix.", "Info", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                    Write-Host $computer "already a member of" $groupWithPrefix
                    Draw_Main_Form
                } else {
                    Add-ADGroupMember -Identity $groupWithPrefix -Members "$computer$"
                    Write-Host $computer "added to" $groupWithPrefix
                    Logging $computer $groupWithPrefix -UseUsernameForLog $true
                    if (-not $GroupsAdded.Contains($groupWithPrefix)) {
                        $GroupsAdded += $groupWithPrefix
                    }
                }
            }
        }

        if ($GroupsAdded.Count -gt 0) {
            if ($UseUsernameForLog -eq $true) {
                Send-Email -GroupNames $GroupsAdded
            } 
        }
        else{
            if ($UseUsernameForLog -eq $true) {
                # Cancel the operation
                Write-Host "Operation canceled by user."
            }   
            Draw_Main_Form
        }
    }  
}



function Send-Email {

    param(

        [string[]]$GroupNames

    )
    $userr = Get-ADUser -Identity $username -Properties Displayname
    $DeployDate = Get-Date -Format "yyyy-MM-dd"
    $Subject = "Service Request: Access Granted to Requested Applications"

    $Styles = ""

    $upperText = @"
Hello $($userr.GivenName), <br/>

<br/>We hope this message finds you well. It is our pleasure to inform you that your requested application(s) which are listed below have been approved and is now in action.<br/>
<br/>To ensure a smooth transition, we kindly ask that you allow for approximately 1 hour for the system to catch up with the recent changes. If you wait for 1 day the application should appear on your computer automatically but if you need it immediately you should be able to find the application(s) in Software Center. If for any reason the applications are not working as expected, we suggest rebooting your computer. If you encounter any further issues or concerns, please do not hesitate to reach out to us for assistance.<br/>

<br/>Instructions to download software(s) from Software Center:<br/>

<br/>

1.     Press Windows button. <br/>

2.     Search 'Software Center'. <br/>

3.     Under applications tab, search for the Application Name. <br/>

4.     Click on Application and click "Install".<br/>

<br/>

"@

    $LowerText = "<BR>If you have any questions, you can reach us at:<BR><BR>By Phone: Helpdesk: 1-844-322-4455 x 2<BR>By Email: FBCTechnicalSupport@FortisBC.com<BR>"

    $tableHeader = @"
<table style="border-width: 1px; border-style: solid; border-color: black; border-collapse: collapse; font-size: 15px;">
    <thead>
        <tr>
            <th style="background: #395870; background: linear-gradient(#49708f, #003865); color: #fff; font-size: 11px; text-transform: uppercase; padding: 10px 15px; vertical-align: middle">Application Name</th>
            <th style="background: #395870; background: linear-gradient(#49708f, #003865); color: #fff; font-size: 11px; text-transform: uppercase; padding: 10px 15px; vertical-align: middle">Deploy Date</th>
        </tr>
    </thead>
    <tbody>
"@

    $tableFooter = @"
    </tbody>
</table>
"@

    $tableRows = ""

    foreach ($group in $GroupNames) {
        $EmailApplication = $group.replace("AP_SCCM_FBC_", "").replace("Required", " ").replace("_", "  ")
        $tableRows += @"
        <tr style="background: #f0f0f2;">
            <td style="border-width: 1px; padding: 3px; border-style: solid; border-color: black;">$EmailApplication</td>
            <td style="border-width: 1px; padding: 3px; border-style: solid; border-color: black;">$DeployDate</td>
        </tr>
"@
    }

    $EmailFrom = "abc@mail.com"
    $EmailTo = "xyz@mail.com"

    $Body = $upperText + $tableHeader + $tableRows + $tableFooter + $LowerText
    $SMTPServer = "smtp.corp.local"
    $SMTPPort = 25

    try {
        Send-Mailmessage -From $EmailFrom -To $EmailTo -Subject $Subject -Body $Body -BodyAsHtml -SmtpServer $SMTPServer -Port $SMTPPort
        Write-Host "Email Sent." -ForegroundColor Green
    } catch [System.Net.Mail.SmtpException] {
        Write-Host "Error: $($Error[0].Exception.InnerException.Message)"
    }

}

function Logging($computerName, $groupName, $UseUsernameForLog = $false) {
    $baseLogFolder = "C:\Users\Public\Desktop\Powershell Scripts\Menu Logs"
    $LoggedInUser = $env:USERNAME

    Write-Host "Creating a Log File..." -ForegroundColor Yellow

    # Create the base log folder if it doesn't exist
    if (-not (Test-Path $baseLogFolder)) {
        New-Item -ItemType Directory -Force -Path $baseLogFolder
    }

    # Determine the folder name based on the UseUsernameForLog flag and username availability
    if ($UseUsernameForLog -and $username) {
        $folderName = $username
    } elseif ($UseUsernameForLog) {
        $baseLogFolder = "C:\Users\Public\Desktop\Powershell Scripts\Menu Logs\~New Device Software Deployment"
        $folderName = $computerName
    }

    $computerLogFolder = Join-Path $baseLogFolder $folderName

    # Create the computer-specific log folder if it doesn't exist
    if (-not (Test-Path $computerLogFolder)) {
        New-Item -ItemType Directory -Force -Path $computerLogFolder
    }

    # Create or append to the log file
    $logFile = "$computerLogFolder\Deployed Softwares.txt"
    $timestamp = Get-Date -Format "DATE: yyyy-MM-dd ||  HH:mm:ss ||"
    $logEntry = "$timestamp - $LoggedInUser added $computerName to $groupName `r`n"
    Add-Content -Path $logFile -Value $logEntry

    Write-Host "Log File Created." -ForegroundColor Green
}

# Add OUs to the ComboBox
$OUApp = "OU=System Center,OU=Groups,OU=Electric,DC=corp,DC=local"
$OUs = @(
    "OU=Fresh PXE Workstations,OU=Workstations,OU=Electric,DC=corp,DC=local",
    "OU=laptops,OU=Workstations,OU=Electric,DC=corp,DC=local",
    "OU=Desktops,OU=Workstations,OU=Electric,DC=corp,DC=local"
    "OU=Stale,OU=Workstations,OU=Electric,DC=corp,DC=local"
)

Draw_Main_Form
