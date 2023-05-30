### PowerShell script for removing a user's groups and disabling the user.
### Add functionality to set "User must change password at next logon"

### Written by Tusshar Singh 2023-4-30
param (
    [string]$username
)

get-date

Add-Type -AssemblyName System.Windows.Forms

function CustomReadHost($prompt) {
    $input = Read-Host -Prompt $prompt

    if ($input -eq 'N') {
        [System.Windows.Forms.MessageBox]::Show("Application Closing. Please Retry Later.", "Operation Cancelled")
        Exit
    }

    return $input
}

$Continue = 'N'

while ($Continue -ne 'Y') {

    Write-Host "============================================"

    $Identity = $username

    $Result = Get-ADUser -Identity $Identity -Properties * | Select Name, SAMAccountName, department, office, LastLogonDate, accountexpirationdate, enabled, PhysicalDeliveryOfficeName, memberof, DistinguishedName

    $date = (Get-Date).tostring("yyyy-MM-dd")

    $name = $Result.name

    Start-Transcript -Path \\trvminfadminp01\c$\ScriptLogs\Closed-$date-$name.txt

    Write-Host "=========Review User Details Below==========="
    Write-Host "Name =" $Result.Name
    Write-Host "Department =" $Result.department
    Write-Host "Office =" $Result.office
    Write-Host "Enabled =" $Result.enabled
    Write-Host "Account Expiration Date =" $Result.accountexpirationdate

    $Continue = CustomReadHost -Prompt 'Is this the correct user?(Y/N): '
}

# Remove all user's groups
$RemoveGroups = CustomReadHost -Prompt 'Attempt to remove all groups for user?(Y/N): '
if ($RemoveGroups -like 'Y'){
    $Groups = $Result.memberof
    foreach ($Group in $Groups){
        try{
            write-host "> Removing group: " $Group
            Remove-ADGroupMember -Identity $Group -Members $Result.SAMAccountName -Confirm:$false
        }
        catch{
            Write-Host "> Failed to remove group: " $Group
        }
    }
}

# Disable User
$DisableUser = CustomReadHost -Prompt 'Disable user?(Y/N): '
if ($DisableUser -like 'Y'){
    if ($Result.enabled -like "True"){
        try{
            Write-Host "> User is currently enabled; Disabling User."
            Disable-ADAccount -identity $Result.SAMAccountName
        }
        catch{
            Write-Host "> Failed to Disable user"
        }
    }
}

# Set "User must change password at next logon"
Write-Host "> Setting 'User must change password at next logon'"
Set-ADUser $Result.SAMAccountName -ChangePasswordAtLogon $true

# Moving AD Object to OU
$MoveUser = CustomReadHost -Prompt 'Move user to Closed OU?(Y/N): '
if ($MoveUser -like 'Y'){
    try{
        $OrganizationalUnit = "OU=Closed,OU=Users,OU=Electric,DC=corp,DC=local"

        Move-ADObject -Identity $Result.DistinguishedName -TargetPath $OrganizationalUnit

        Write-Host "> Moved user to Closed OU"

    }
    catch{

        Write-Host "> Failed to move user to Closed OU"

    }
}

# Set account expiration date to the current date
$SetAccountExpiry = CustomReadHost -Prompt 'Set account expiration date to the current date?(Y/N): '
if ($SetAccountExpiry -like 'Y') {
    try {
        $CurrentDate = Get-Date -Hour 23 -Minute 59 -Second 59
        Set-ADUser -Identity $Result.SAMAccountName -AccountExpirationDate $CurrentDate
        Write-Host "> Set account expiration date to the current date"
    }
    catch {
        Write-Host "> Failed to set account expiration date to the current date"
    }
}


# Rename User to User (Closed month dd, yyyy)
$RenameUser = CustomReadHost -Prompt 'Rename user with close date?(Y/N): '
if ($RenameUser -like 'Y'){
    try{
        $Result = Get-ADUser -Identity $Identity -Properties * | Select DistinguishedName, Name
        $CloseDate = Read-Host -Prompt 'Enter Date Format (Month Day, YYYY) ex September 1, 2018: '
        $UserRename = $Result.Name + " (Closed, " + $CloseDate + ")"
        Rename-ADObject -Identity $Result.DistinguishedName -NewName $UserRename
        Write-Host "> Renamed to:" $UserRename
    }
    catch{
        Write-Host "> Rename Failed"
    }
} 

Stop-Transcript