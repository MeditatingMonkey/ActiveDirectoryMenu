Add-Type -AssemblyName System.Windows.Forms

    $detailsForm = New-Object System.Windows.Forms.Form
    $detailsForm.Text = "User Details: $username"
    $detailsForm.Size = New-Object System.Drawing.Size(600, 500)
    $detailsForm.StartPosition = "CenterScreen"

    $propertiesToFetch = @('GivenName', 'Surname', 'SamAccountName', 'Title', 'Description', 'Company', 'OfficePhone', 'Office', 'WhenCreated', 'AccountExpirationDate', 'PasswordLastSet', 'Manager')
    $userDetails = Get-ADUser -Identity $username -Properties $propertiesToFetch

    $y = 10
    $textBoxes = @{}
    foreach ($property in $propertiesToFetch) {
        if ($property -notin @('WhenCreated', 'AccountExpirationDate', 'PasswordLastSet', 'Manager')) {

            $label = New-Object System.Windows.Forms.Label
            $label.Location = New-Object System.Drawing.Point(10, $y)
            $label.AutoSize = $true
            $label.Text = "$($property):"
            $label.Font = New-Object System.Drawing.Font('Verdana Pro',9,[System.Drawing.Fontstyle]::Bold)
            $detailsForm.Controls.Add($label)

            if ($property -eq 'SamAccountName') {
            $label.Text = "Username:"
            } 
            else {
                $label.Text = "$($property):"
            }

            $textBox = New-Object System.Windows.Forms.TextBox
            $textBox.Location = New-Object System.Drawing.Point(200, ($y - 3))
            $textBox.Size = New-Object System.Drawing.Size(350, 20)
            $textBox.Text = $userDetails.$property
            $textBox.Font = New-Object System.Drawing.Font('Verdana Pro',10)
            $detailsForm.Controls.Add($textBox)

            $textBoxes[$property] = $textBox
            $y += 30
        }
    }


    foreach ($property in @('WhenCreated', 'AccountExpirationDate', 'PasswordLastSet')) {
        $label = New-Object System.Windows.Forms.Label
        $label.Location = New-Object System.Drawing.Point(10, $y)
        $label.AutoSize = $true
        $label.Font = New-Object System.Drawing.Font('Verdana Pro',9)

        if ($property -eq 'PasswordLastSet') 
        {
            $daysSincePasswordLastSet = ((Get-Date) - $userDetails.PasswordLastSet).Days 
            $label.Text = "Password Last Set:  ($daysSincePasswordLastSet Days Ago)"
        } 
        
        Else
        { 
            $label.Text = "$($property):  $($userDetails.$property)"
        }

        $detailsForm.Controls.Add($label)
        $y += 30
    }

    if ($userDetails.Manager) {
        $managerDetails = Get-ADUser -Identity $userDetails.Manager -Properties Name, Department
        $managerName = $managerDetails.Name
        $managerDepartment = $managerDetails.Department
    } 
    else {
        $managerName = "N/A"
        $managerDepartment = "N/A"
    }

    $managerlabel = New-Object System.Windows.Forms.Label
    $managerlabel.Text = "Manager name:"
    $managerlabel.Location = New-Object System.Drawing.Point(10, $y)
    $managerlabel.AutoSize = $true
    $managerlabel.Font = New-Object System.Drawing.Font('Verdana Pro',9,[System.Drawing.Fontstyle]::Bold)
    $detailsForm.Controls.Add($managerlabel)

    $ManagertextBox = New-Object System.Windows.Forms.Textbox
    $ManagertextBox.Location = New-Object System.Drawing.Point(200, $y)
    $ManagertextBox.Size = New-Object System.Drawing.Size(250, 20)
    $ManagertextBox.Text = $managerName
    $ManagerTextBox.Font = 'Verdana Pro,10'
    $detailsForm.Controls.Add($ManagertextBox)

    $CheckButton = New-Object System.Windows.Forms.Button
    $CheckButton.Location = New-Object System.Drawing.Point(460, $y)
    $CheckButton.Size = New-Object System.Drawing.Size(90, 24)
    $CheckButton.Text = "Check"
    $CheckButton.Add_Click({
    $inputManagerName = $ManagertextBox.Text
    $managerDetailsList = Get-ADUser -Filter "Name -like '*$inputManagerName*'" -Properties Name, Department

    if ($managerDetailsList) {
        $selectionForm = New-Object System.Windows.Forms.Form
        $selectionForm.Text = "Select Manager"
        $selectionForm.Size = New-Object System.Drawing.Size(400, 300)
        $selectionForm.StartPosition = "CenterScreen"

        $listBox = New-Object System.Windows.Forms.ListBox
        $listBox.Location = New-Object System.Drawing.Point(10, 10)
        $listBox.Size = New-Object System.Drawing.Size(360, 200)
        $listBox.DisplayMember = 'Name'

        foreach ($managerDetails in $managerDetailsList) {
            $listBox.Items.Add($managerDetails)
        }

        $selectionForm.Controls.Add($listBox)

        $selectButton = New-Object System.Windows.Forms.Button
        $selectButton.Location = New-Object System.Drawing.Point(50, 220)
        $selectButton.Size = New-Object System.Drawing.Size(100, 30)
        $selectButton.Text = "Select"
        $selectButton.Add_Click({
            if ($listBox.SelectedItem -ne $null) {
                $selectedManager = $listBox.SelectedItem
                $ManagertextBox.Text = $selectedManager.Name
                $departmentValueLabel.Text = $selectedManager.Department
                $selectionForm.Close()
            }
        })
        $selectionForm.Controls.Add($selectButton)

        $cancelButton = New-Object System.Windows.Forms.Button
        $cancelButton.Location = New-Object System.Drawing.Point(250, 220)
        $cancelButton.Size = New-Object System.Drawing.Size(100, 30)
        $cancelButton.Text = "Cancel"
        $cancelButton.Add_Click({ $selectionForm.Close() })
        $selectionForm.Controls.Add($cancelButton)

        $selectionForm.ShowDialog()
    } else {
        [System.Windows.Forms.MessageBox]::Show("Invalid manager name. Please enter a valid manager name.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})
    $detailsForm.Controls.Add($CheckButton)

    $y += 30

    $departmentLabel = New-Object System.Windows.Forms.Label
    $departmentLabel.Text = "Manager Department:"
    $departmentLabel.Location = New-Object System.Drawing.Point(10, $y)
    $departmentLabel.AutoSize = $true
    $departmentLabel.Font = New-Object System.Drawing.Font('Verdana Pro', 9,[System.Drawing.Fontstyle]::Bold)
    $detailsForm.Controls.Add($departmentLabel)

    $departmentValueLabel = New-Object System.Windows.Forms.Label
    $departmentValueLabel.Text = $managerDepartment
    $departmentValueLabel.Location = New-Object System.Drawing.Point(200, $y)
    $departmentValueLabel.AutoSize = $true
    $departmentValueLabel.Font = New-Object System.Drawing.Font('Verdana Pro', 9)
    $detailsForm.Controls.Add($departmentValueLabel)

    $saveButton = New-Object System.Windows.Forms.Button
    $saveButton.Location = New-Object System.Drawing.Point(15, 410)
    $saveButton.Size = New-Object System.Drawing.Size(550, 30)
    $saveButton.Text = "Save"
  
    $saveButton.Add_Click({

        $newValues = @{}
        $changes = @()
        foreach ($property in $propertiesToFetch) 
        {
            if ($property -notin @('WhenCreated', 'AccountExpirationDate', 'PasswordLastSet', 'Manager')) {
                $oldValue = $userDetails[$property]
                $newValue = $textBoxes[$property].Text
                if ($oldValue -ne $newValue) {
                    $newValues[$property] = $newValue
                    $changes += "$property`: $oldValue -> $newValue"
                }
            }
        }

        # Check if the manager attribute has changed
        $oldManager = $userDetails['Manager']
        $newManager = Get-ADUser -Filter "Name -eq '$($ManagertextBox.Text)'" -Properties Name
        if ($oldManager -ne $newManager.DistinguishedName) {
            $oldManagerName = (Get-ADUser -Identity $oldManager[0] -Properties Name).Name

            $newValues['Manager'] = $newManager.DistinguishedName
            $changes += "Manager`: $oldManagerName -> $($newManager.Name)"
        }

        if ($changes) {
            $changesString = $changes -join "`r`n"
            $messageBoxResult = [System.Windows.Forms.MessageBox]::Show("Please confirm the changes:`n`n$changesString", "Confirm Changes", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)

            if ($messageBoxResult -eq 'Yes') {
                try {
                    Set-ADUser -Identity $username @newValues
                    [System.Windows.Forms.MessageBox]::Show("User details have been updated.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                    $detailsForm.Close()
                    Logging -NewValues $newValues -LogFile "$username User Details.txt"
                }
                catch {
                    [System.Windows.Forms.MessageBox]::Show("An error occurred while updating the user details.`n`n$($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                }
            }
        } else {
            [System.Windows.Forms.MessageBox]::Show("No changes were made.", "No Changes", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
    })

function Logging {
    param (
        [hashtable]$NewValues,
        [string]$LogFile = "$username User Details.txt"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd || HH:mm:ss"
    $currentUser = $env:USERNAME
    $logFolderPath = "C:\Users\Public\Desktop\Powershell Scripts\Menu Logs\$username"
    $logFilePath = Join-Path $logFolderPath $LogFile

    if (-not (Test-Path $logFolderPath)) {
        New-Item -ItemType Directory -Path $logFolderPath
    }

    $logMessage = "Admin: $currentUser`r`n"
    $logMessage += "Made these changes to $username : [$timestamp]`r`n********************************************************************************`r`n"

    foreach ($key in $NewValues.Keys) {
        $logMessage += "`t$key`: $($NewValues[$key])`r`n"
    }

    $logMessage += "`r`n********************************************************************************`r`n"

    try {
        Add-Content -Path $logFilePath -Value $logMessage
    }
    catch {
        Write-Host "Failed to write log: $($_.Exception.Message)"
    }
}

   $detailsForm.Controls.Add($saveButton)

   $detailsForm.ShowDialog()