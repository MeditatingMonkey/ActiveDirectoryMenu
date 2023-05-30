# Import required modules
Import-Module ActiveDirectory
Add-Type -AssemblyName System.Windows.Forms

$userr = Get-ADUser -Identity $username -Properties DisplayName
$DisplayName = $userr.DisplayName

# Create form
$form = New-Object System.Windows.forms.form
$form.text = "Password Reset - $DisplayName"
$form.Size = New-Object System.Drawing.Size(400, 200)
$form.StartPosition = "CenterScreen"

$labelNewPassword = New-Object System.Windows.Forms.Label
$labelNewPassword.Text = "New Password:"
$labelNewPassword.Location = New-Object System.Drawing.Point(10, 20)
$labelNewPassword.Font = New-Object System.Drawing.Font('Verdana Pro',10)

$textboxNewPassword = New-Object System.Windows.Forms.TextBox
$textboxNewPassword.Location = New-Object System.Drawing.Point(150, 15)
$textboxNewPassword.Size = New-Object System.Drawing.Size(200, 20)
$textboxNewPassword.UseSystemPasswordChar = $true
$textboxNewPassword.Font = New-Object System.Drawing.Font('Verdana Pro',10)

$labelconfirmPassword = New-Object System.Windows.Forms.Label
$labelconfirmPassword.Text = "Confirm Password:"
$labelconfirmPassword.Location = New-Object System.Drawing.Point(10, 50)
$labelconfirmPassword.AutoSize = $true
$labelconfirmPassword.Font = New-Object System.Drawing.Font('Verdana Pro',10)

$textboxconfirmPassword = New-Object System.Windows.Forms.TextBox
$textboxconfirmPassword.Location = New-Object System.Drawing.Point(150, 45)
$textboxconfirmPassword.Size = New-Object System.Drawing.Size(200, 20)
$textboxconfirmPassword.Font = New-Object System.Drawing.Font('Verdana Pro',10)
$textboxconfirmPassword.UseSystemPasswordChar = $true

$checkbox = New-Object System.Windows.Forms.CheckBox
$checkbox.Text = "Show Password"
$checkbox.Location = New-Object System.Drawing.Point(200,75)

$checkbox.Add_CheckedChanged({
    if($checkbox.Checked) {
        $textboxNewPassword.UseSystemPasswordChar = $false
        $textboxconfirmPassword.UseSystemPasswordChar = $false
    }
    else {
        $textboxNewPassword.UseSystemPasswordChar = $true
        $textboxconfirmPassword.UseSystemPasswordChar = $true
    }
})

$form.Controls.Add($checkbox)

$buttonResetPassword = New-Object System.Windows.Forms.Button
$buttonResetPassword.Text = "Reset Password"
$buttonResetPassword.Location = New-Object System.Drawing.Point(150, 105)
$buttonResetPassword.Size = New-Object System.Drawing.Size(200, 35)
$buttonResetPassword.Font = New-Object System.Drawing.Font('Verdana Pro',9)

$form.Controls.Add($labelNewPassword)
$form.Controls.Add($textboxNewPassword)
$form.Controls.Add($buttonResetPassword)
$form.Controls.Add($labelconfirmPassword)
$form.Controls.Add($textboxconfirmPassword)

# Event handler for button click
$buttonResetPassword.Add_Click({
    $UserSAM = $username
    $NewPasswordPlainText = $textboxNewPassword.Text
    $NewPassword = ConvertTo-SecureString -String $NewPasswordPlainText -AsPlainText -Force

    if ($textboxNewPassword.Text -eq $textboxconfirmPassword.Text) {
        # Get the default domain password policy
        $PasswordPolicy = Get-ADDefaultDomainPasswordPolicy

        # Check if the new password meets the policy requirements
        $meetsLengthRequirement = $NewPasswordPlainText.Length -ge $PasswordPolicy.MinPasswordLength
        $meetsComplexityRequirement = $PasswordPolicy.ComplexityEnabled -eq $false -or (([Regex]::Matches($NewPasswordPlainText, "[A-Z]").Count -gt 0) -and ([Regex]::Matches($NewPasswordPlainText, "[a-z]").Count -gt 0) -and ([Regex]::Matches($NewPasswordPlainText, "[0-9]").Count -gt 0))

        if ($meetsLengthRequirement -and $meetsComplexityRequirement) {
            # Reset the user's password
            Set-ADAccountPassword -Identity $UserSAM -NewPassword $NewPassword -Reset

            # OPTIONAL: Force user to change password at next logon
            Set-ADUser -Identity $UserSAM -ChangePasswordAtLogon $true

            [System.Windows.Forms.MessageBox]::Show("Password has been reset successfully!", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            $form.Dispose()
        } 
        else {
        [System.Windows.Forms.MessageBox]::Show("The new password does not meet the domain password policy requirements.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    } 
    else {
        [System.Windows.Forms.MessageBox]::Show("The Passwords do not match.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }    
})

#Show form
[void]$form.ShowDialog()