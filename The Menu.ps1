##################################################
##
##My Active Directory
##
##Description:
##This PowerShell Script is a GUI representation for tasks like Software Deployment, Group Accesses, Unlocking Accounts
##Editing Details, Closing a User, Deploying a software to a new computer and more.
##
##By : Tusshar Singh
##
##Date: 24 April 2023
##
##################################################

Add-Type -AssemblyName System.Windows.Forms
import-module ActiveDirectory

Add-Type -AssemblyName PresentationFramework

$SearchBaseOU = "OU=Users,OU=Electric,DC=corp,DC=local"

function ShowMenu($menuForm){

    $Button1 = New-Object $BtnObject
    $Button1.Location = New-Object $Point(120, 10)
    $Button1.Size = New-Object $Size(150,50)
    $Button1.Text = "UNLOCK ACCOUNT"
    $Button1.Font = New-Object System.Drawing.Font('Verdana Pro',9)
    $menuForm.Controls.Add($Button1)
    $Button1.Add_Click({Unlock-ADAccount -Identity $username;$menuForm.Dispose();$menuForm.Update();Write-Host "Account Unlocked."})

    $Button2 = New-Object $BtnObject
    $Button2.Location = New-Object $Point(120, 80)
    $Button2.Size = New-Object $Size(150,50)
    $Button2.Text = "REMOTE ACCESS"
    $Button2.Font = New-Object System.Drawing.Font('Verdana Pro',9)
    $menuForm.Controls.Add($Button2)
    $Button2.Add_Click({RemoteAccess})

    $Button3 = New-Object $BtnObject
    $Button3.Location = New-Object $Point(120, 150)
    $Button3.Size = New-Object $Size(150,50)
    $Button3.Text = "SOFTWARE DEPLOYMENT"
    $Button3.Font = New-Object System.Drawing.Font('Verdana Pro',9)
    $menuForm.Controls.Add($Button3)
    $Button3.Add_Click({SoftwareDeployment -selectedIndex 4 -addPrimaryDevice $true -useUsernameForLog $true})

    $Button4 = New-Object $BtnObject
    $Button4.Location = New-Object $Point(120, 220)
    $Button4.Size = New-Object $Size(350,50)
    $Button4.Text = "USER DETAILS"
    $Button4.Font = New-Object System.Drawing.Font('Verdana Pro',9)
    $menuForm.Controls.Add($Button4)
    $Button4.Add_Click({ShowUserDetails})

    $Button5 = New-Object $BtnObject
    $Button5.Location = New-Object $Point(320, 10)
    $Button5.Size = New-Object $Size(150,50)
    $Button5.Text = "CLOSE USER"
    $Button5.Font = New-Object System.Drawing.Font('Verdana Pro',9)
    $menuForm.Controls.Add($Button5)
    $Button5.Add_Click({CloseUser})
            
    $Button6 = New-Object $BtnObject
    $Button6.Location = New-Object $Point(320, 80)
    $Button6.Size = New-Object $Size(150,50)
    $Button6.Text = "RESET PASSWORD"
    $Button6.Font = New-Object System.Drawing.Font('Verdana Pro',9)
    $menuForm.Controls.Add($Button6)
    $Button6.Add_Click({ResetPassword})

    $Button7 = New-Object $BtnObject
    $Button7.Location = New-Object $Point(320, 150)
    $Button7.Size = New-Object $Size(150,50)
    $Button7.Text = "GROUP ACCESS"
    $Button7.Font = New-Object System.Drawing.Font('Verdana Pro',9)
    $menuForm.Controls.Add($Button7)
    $Button7.Add_Click({GroupDeployment})
    
    # Add a label for the user lock status
    $StatusLabel = New-Object $LabelObject
    $StatusLabel.Location = New-Object $Point(126, 63)
    $StatusLabel.AutoSize = $true

    # Check if the user account is locked
    $user = Get-ADUser -Identity $username -Properties LockedOut, Displayname
    if ($user.LockedOut) {
        $StatusLabel.Text = "  User is currently locked"
        $StatusLabel.ForeColor = [System.Drawing.Color]::Red
    } else {
        $StatusLabel.Text = "User is currently unlocked"
        $StatusLabel.ForeColor = [System.Drawing.Color]::Green
    }

    $menuForm.Controls.Add($StatusLabel)

}

function RemoteAccess(){
$scriptPath = 'C:\Users\Public\Desktop\Powershell Scripts\Upcoming\Remote Access.ps1'
    if (Test-Path $scriptPath)
        {
            . $scriptPath
        } 
    else{
            Write-Error "Could not find script at path $scriptPath"
        }
}

function SoftwareDeployment($selectedIndex, $addPrimaryDevice, $useUsernameForLog = $false){
$scriptPath = 'C:\Users\Public\Desktop\Powershell Scripts\Upcoming\Application Deployment.ps1'
    if (Test-Path $scriptPath)
        {
            . $scriptPath -SelectedIndex $selectedIndex -AddPrimaryDevice $addPrimaryDevice -UseUsernameForLog $useUsernameForLog
        } 
    else{
            Write-Error "Could not find script at path $scriptPath"
        }
}  

function ShowUserDetails(){
$scriptPath = 'C:\Users\Public\Desktop\Powershell Scripts\Upcoming\Show User Details.ps1'
    if (Test-Path $scriptPath) 
        {
            . $scriptPath
        } 
    
    else{
            Write-Error "Could not find script at path $scriptPath"
        }
}

Function CloseUser(){
    if ($Userlist.SelectedItem -ne $null)
    {
        # Use the correct variable for the selected user
        $displayName = $Userlist.SelectedItem.ToString()
        $username = $UserMapping[$displayName]

        $scriptPath = "C:\Users\Public\Desktop\Powershell Scripts\Upcoming\CloseUser.ps1"
        $arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`" -username $username"
        Start-Process -FilePath "powershell.exe" -ArgumentList $arguments -WindowStyle Normal
    }
    else
    {
        [System.Windows.Forms.MessageBox]::Show("Please select a user from the list.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

function GroupDeployment(){
$scriptPath = 'C:\Users\Public\Desktop\Powershell Scripts\Upcoming\Group Deployment.ps1'
    if (Test-Path $scriptPath) 
        {
            . $scriptPath
        } 
    
    else{
            Write-Error "Could not find script at path $scriptPath"
        }
}

function ResetPassword{
$scriptPath = 'C:\Users\Public\Desktop\Powershell Scripts\Upcoming\Reset Password.ps1'
    if (Test-Path $scriptPath) 
        {
            . $scriptPath
        } 
    
    else{
            Write-Error "Could not find script at path $scriptPath"
        }
}


$FormObject = [System.Windows.Forms.Form]
$LabelObject = [System.Windows.Forms.Label]
$TxtObject = [System.Windows.Forms.TextBox]
$BtnObject = [System.Windows.Forms.Button]
$Point = [System.Drawing.Point]
$Size = [System.Drawing.Size]

$Form1 = New-Object $FormObject
$Form1.Text = "My Active Directory"
$Form1.Width = 585
$Form1.Height = 280
$Form1.StartPosition = "CenterScreen"

# Create a label for the username
$Label = New-Object $LabelObject
$Label.Location = New-Object $Point(55,10)
$Label.AutoSize = $true
$Label.Text = "NAME :"
$label.Font = New-Object System.Drawing.Font('Verdana Pro',9,[System.Drawing.Fontstyle]::Bold)
$Form1.Controls.Add($Label)

# Create a label for the username
$Label2 = New-Object $LabelObject
$Label2.Location = New-Object $Point(30,40)
$Label2.AutoSize = $true
$Label2.Text = "USER LIST  :"
$label2.Font = New-Object System.Drawing.Font('Verdana Pro',9)
$Form1.Controls.Add($Label2)

# Create a text box for the username
$TextBox = New-Object $TxtObject
$TextBox.Location = New-Object $Point (110, 6)
$TextBox.Size = New-Object $Size(300)
$TextBox.Font = 'Verdana Pro,11'
$Form1.Controls.Add($TextBox)

$User = Get-ADUser -Filter * -Properties Name, Displayname, SamAccountName -SearchBase $SearchBaseOU 
$UserMapping = @{}

$Userlist = New-Object System.Windows.Forms.ListBox
$Userlist.Location = New-Object $Point (110, 40)
$Userlist.Size = New-Object $Size(300,180)
$Userlist.Font = 'Verdana Pro,11'
$Form1.Controls.Add($Userlist)

Function UpdateUser{
        $SearchUser = $TextBox.Text
        $WildcardSearchUser = "*$SearchUser*"
        $FilteredUsers = $User | Where-Object { ($_.Name -like $WildcardSearchUser) } #String approximation
        $Userlist.Items.Clear() #refreshing the list

    foreach ($FilteredUser in $FilteredUsers) {
        $Userlist.Items.Add($FilteredUser.DisplayName)
        $UserMapping[$FilteredUser.DisplayName] = $FilteredUser.SamAccountName 
    }
}

$TextBox.Add_TextChanged({
    UpdateUser
})

# Create a button to unlock the user account
$Button = New-Object $BtnObject
$Button.Location = New-Object $Point(417, 5)
$Button.Size = New-Object $Size(150,26)
$Button.Text = "SUBMIT"
$Button.Font = New-Object System.Drawing.Font('Verdana Pro',9,[System.Drawing.FontStyle]::Bold)
$Button.Add_Click({
    if ($Userlist.SelectedItem -ne $null)
    {
        #Use the correct variable fo rthe selected user
        $displayName = $Userlist.SelectedItem.ToString()
        $username = $UserMapping[$displayName]

        $menuForm = New-Object System.Windows.Forms.Form
        $menuForm.Text = Get-ADUser -Identity $username -Properties DisplayName | Select-Object -ExpandProperty DisplayName
        $menuForm.Size = New-Object $Size(600,400)
        $menuForm.StartPosition = "CenterScreen"

        ShowMenu -menuForm $menuForm

        $menuForm.ShowDialog()
    }
    else
    {
        [System.Windows.Forms.MessageBox]::Show("User does not exist in Active Directory.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
    }
})



$Form1.Controls.Add($Button)


$Button8 = New-Object $BtnObject
$Button8.Location = New-Object $Point(417, 40)
$Button8.Size = New-Object $Size(150,40)
$Button8.Text = "NEW USER"
$Button8.Font = New-Object System.Drawing.Font('Verdana Pro',9)

$Form1.Controls.Add($Button8)

$Button9 = New-Object $BtnObject
$Button9.Location = New-Object $Point(417, 85)
$Button9.Size = New-Object $Size(150,40)
$Button9.Text = "SOFTWARE DEPLOYEMENT"
$Button9.Font = New-Object System.Drawing.Font('Verdana Pro',9)
$Button9.Add_Click({SoftwareDeployment -selectedIndex 0 -addPrimaryDevice $false})
$Form1.Controls.Add($Button9)

$Form1.ShowDialog() | Out-Null