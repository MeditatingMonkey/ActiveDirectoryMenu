# Load necessary assemblies
Add-Type -AssemblyName System.Windows.Forms

# Create form
$form = New-Object System.Windows.Forms.Form
$form.Size = New-Object System.Drawing.Size(400, 150)
$form.Text = 'Enter Computer Number'
$form.StartPosition = "CenterScreen"

# Create label
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(20, 20)
$label.Size = New-Object System.Drawing.Size(120, 20)
$label.Text = 'Computer Number:'
$form.Controls.Add($label)

# Create input field
$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(150, 20)
$textBox.Size = New-Object System.Drawing.Size(200, 20)
$form.Controls.Add($textBox)

# Create OK button
$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(150, 60)
$okButton.Size = New-Object System.Drawing.Size(75, 23)
$okButton.Text = 'OK'
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.Controls.Add($okButton)

# Create Cancel button
$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Location = New-Object System.Drawing.Point(275, 60)
$cancelButton.Size = New-Object System.Drawing.Size(75, 23)
$cancelButton.Text = 'Cancel'
$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.Controls.Add($cancelButton)

# Show form and get result
$result = $form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
    $computerNumber = $textBox.Text

    $SiteCode = 'FEI'
    $SCCMPSModule = 'C:\Program Files (x86)\Configuration Manager\Console\bin\ConfigurationManager\ConfigurationManager.psd1'

    Import-Module -Name $SCCMPSModule
    Set-Location "$($SiteCode):"

    try {
        Invoke-CMRemoteControl -DeviceName $computerNumber
    } catch [System.Management.Automation.ItemNotFoundException] {
        [System.Windows.Forms.MessageBox]::Show("Computer not found. Please check the computer number and try again.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    } catch {
        [System.Windows.Forms.MessageBox]::Show("An unexpected error occurred. Please try again.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}
