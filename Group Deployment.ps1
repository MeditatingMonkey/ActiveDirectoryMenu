Import-Module ActiveDirectory

# Load the necessary assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Define the form
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Active Directory Groups'
$form.Size = New-Object System.Drawing.Size(500, 535)
$form.StartPosition = 'CenterScreen'

$SearchBoxLabel = New-Object System.Windows.Forms.Label
$SearchBoxLabel.Text = 'Search Group:'
$SearchBoxLabel.AutoSize = $true
$SearchBoxLabel.Font = New-Object System.Drawing.Font('Verdana Pro',10)
$SearchBoxLabel.Location = New-Object System.Drawing.Point(10, 5)

# Define the search textbox
$searchBox = New-Object System.Windows.Forms.TextBox
$searchBox.Location = New-Object System.Drawing.Point(10, 28)
$searchBox.Size = New-Object System.Drawing.Size(460, 20)
$searchBox.Font = New-Object System.Drawing.Font('Verdana Pro',10)

# Define the listbox for groups
$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Location = New-Object System.Drawing.Point(10, 50)
$listBox.Size = New-Object System.Drawing.Size(460, 280)
$listBox.Font = New-Object System.Drawing.Font('Verdana Pro',10)

# Get the groups in the specified OU and add them to the listbox
$groups = Get-ADGroup -Filter * -SearchBase "OU=Groups,OU=Electric,DC=corp,DC=local"
foreach ($group in $groups) {
    $listBox.Items.Add($group.Name)
}

$selectedGroupsLabel = New-Object System.Windows.Forms.Label
$selectedGroupsLabel.Text = 'Selected Group (s):'
$selectedGroupsLabel.AutoSize = $true
$selectedGroupsLabel.Font = New-Object System.Drawing.Font('Verdana Pro',10)
$selectedGroupsLabel.Location = New-Object System.Drawing.Point(10, 330)

# Define the listbox to display the selected groups
$selectedGroupsListBox = New-Object System.Windows.Forms.ListBox
$selectedGroupsListBox.Location = New-Object System.Drawing.Point(10, 350)
$selectedGroupsListBox.Size = New-Object System.Drawing.Size(460, 100)
$selectedGroupsListBox.Font = New-Object System.Drawing.Font('Verdana Pro',10)
$selectedGroupsListBox.SelectionMode = 'MultiExtended'

# Store the selected group names
$selectedGroupNames = New-Object 'System.Collections.Generic.HashSet[string]'

# Update the selected group names list when the ListBox selection changes
$listBox.Add_SelectedIndexChanged({
    foreach ($selectedGroup in $listBox.SelectedItems) {
        if ($selectedGroupNames.Add($selectedGroup)) {
            $selectedGroupsListBox.Items.Add($selectedGroup)
        }
    }
})

# Search function for the groups
$searchBox.Add_TextChanged({
    $listBox.Items.Clear()
    $searchText = $searchBox.Text
    $filteredGroups = $groups | Where-Object { $_.Name -like "*$searchText*" } 
    foreach ($group in $filteredGroups) {
        $listBox.Items.Add($group.Name)
    }
})

# Define the Clear Selection button
$clearSelectionButton = New-Object System.Windows.Forms.Button
$clearSelectionButton.Location = New-Object System.Drawing.Point(10, 460)
$clearSelectionButton.Size = New-Object System.Drawing.Size(100, 30)
$clearSelectionButton.Text = 'Clear Selection'
$clearSelectionButton.Add_Click({
    foreach ($selectedGroup in $selectedGroupsListBox.SelectedItems) {
        $selectedGroupNames.Remove($selectedGroup)
    }
    $selectedGroupsListBox.Items.Clear()
    foreach ($selectedGroupName in $selectedGroupNames) {
        $selectedGroupsListBox.Items.Add($selectedGroupName)
    }
})

# Define the Clear All button
$clearAllButton = New-Object System.Windows.Forms.Button
$clearAllButton.Location = New-Object System.Drawing.Point(120, 460)
$clearAllButton.Size = New-Object System.Drawing.Size(100, 30)
$clearAllButton.Text = 'Clear All'
$clearAllButton.Add_Click({
    $selectedGroupNames.Clear()
    $selectedGroupsListBox.Items.Clear()
})

function Write-Log {
    param(
        [string]$Username,
        [string[]]$GroupNames
    )

    $LoggedInUser = $env:USERNAME
    $userFolder = $Username

    Write-Host "Creating a Log File." -ForegroundColor "Yellow"

    if (-not (Test-Path -Path "C:\Users\Public\Desktop\Powershell Scripts\Menu Logs\$userFolder")) {
        New-Item -ItemType Directory -Path "C:\Users\Public\Desktop\Powershell Scripts\Menu Logs\$userFolder" -ErrorAction SilentlyContinue
    }

    $logFilePath = "C:\Users\Public\Desktop\Powershell Scripts\Menu Logs\$userFolder\$Username Group Access.txt"   
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] - $LoggedInUser Added user $Username to the following groups:`n"
    $logEntry += "****************************************************************`n"

    foreach ($groupName in $GroupNames) {
        $logEntry += "`t$groupName`n"
    }

    $logEntry += "****************************************************************`n"

    try {
        Add-Content -Path $logFilePath -Value $logEntry -ErrorAction Stop
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to write to log file. Error: $($_.Exception.Message)", "Error", [Syste.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Hand)
    }

    Write-Host "Done" -ForegroundColor "Green"
}


# Define the OK button
$button = New-Object System.Windows.Forms.Button
$button.Location = New-Object System.Drawing.Point(370, 460)
$button.Size = New-Object System.Drawing.Size(100, 30)
$button.Text = 'OK'
$button.DialogResult = [System.Windows.Forms.DialogResult]::OK
$button.Add_Click({

    $confirmationMessage = "Are you sure you want to add user $username to the following groups?" + [Environment]::NewLine
    foreach ($selectedGroup in $selectedGroupNames) {
        $confirmationMessage += "`n$selectedGroup"
    }

    $confirmationResult = [System.Windows.Forms.MessageBox]::Show($confirmationMessage, "Confirmation", [System.Windows.Forms.MessageBoxButtons]::OKCancel)

    if ($confirmationResult -eq [System.Windows.Forms.DialogResult]::OK) {
        $addedGroups = @()

        foreach ($selectedGroup in $selectedGroupNames) {
            try {
                Add-ADGroupMember -Identity $selectedGroup -Members $username -ErrorAction Stop
                $addedGroups += $selectedGroup
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Failed to add user $username to group $selectedGroup. Error: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Hand)

                return
            }
        }
        # Call the logging function outside the loop
        Write-Log -Username $username -GroupNames $addedGroups

        if ($addedGroups.Count -gt 0) {
            Send-Email -GroupNames $addedGroups
        } else {
            # Cancel the operation
            Write-Host "Operation canceled by user."
            Draw_Main_Form
        }
        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK

        $form.Close()
    }
})


function Send-Email {

    param(
        [string[]]$GroupNames
    )

    $userr = Get-ADUser -Identity $username -Properties Displayname
    $DeployDate = Get-Date -Format "yyyy-MM-dd"
    $Subject = "Service Request: Access Granted to Requested Groups"

    $Styles = ""

    $upperText = @"
Hello $($userr.GivenName), <br/>

<br/>We hope this message finds you well. We are writing to inform you about the status of your recent service request for the Group(s) listed below. <br/>

<br/>It is our pleasure to inform you that your request has been approved and is now in action. To ensure a smooth transition, we kindly ask that you allow for approximately 1 hour for the system to catch up with the recent changes. If for any reason you are not able to access the groups, we suggest rebooting your computer. If you encounter any further issues or concerns, please do not hesitate to reach out to us for assistance.<br/>

<br/>

"@

    $LowerText = "<BR>If you have any questions, you can reach us at:<BR><BR>By Phone: Helpdesk: 1-844-322-4455 x 2<BR>By Email: FBCTechnicalSupport@FortisBC.com<BR>"

    $tableHeader = @"
<table style="border-width: 1px; border-style: solid; border-color: black; border-collapse: collapse; font-size: 15px;">
    <thead>
        <tr>
            <th style="background: #395870; background: linear-gradient(#49708f, #003865); color: #fff; font-size: 11px; text-transform: uppercase; padding: 10px 15px; vertical-align: middle">Folder Name</th>
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
        $EmailGroup = $group.replace("BCTR_", "").replace("FA_", "").replace("BC_", "").replace("_", "  ")
        $tableRows += @"
        <tr style="background: #f0f0f2;">
            <td style="border-width: 1px; padding: 3px; border-style: solid; border-color: black;">$EmailGroup</td>
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

$form.AcceptButton = $button

$form.Controls.AddRange(@($SearchBoxLabel,$SearchBox,$listBox,$selectedGroupsLabel,$selectedGroupsListBox,$clearSelectionButton,$clearAllButton, $button))
# Show the form and get the result
$result = $form.ShowDialog()
