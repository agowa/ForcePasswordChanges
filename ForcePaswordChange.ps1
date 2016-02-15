## ========================== Settings ====================================

[System.String] $ou = "OU=Department,DC=contoso,DC=com";
[System.DateTime] $pwdlastChanged = (Get-Date).AddYears(-5);
##                                                     ^^^^^
## ======================== END Settings ==================================

## Change password on next Login => pwdLastSet -eq 0
## Where-Object pwdLastSet -NE 0   ==>>> Where User does not need to change password at next login
## PasswordLastSet => Human readable pwdLastSet
$UsersWhichMustChangePasswords = Get-ADUser -SearchBase $ou -filter * -properties pwdLastSet,PasswordLastSet,PasswordNeverExpires,LastLogonDate | Where-Object Enabled -EQ $True | Where-Object PasswordNeverExpires -EQ $False | Where-Object pwdLastSet -NE 0 | Where-Object PasswordLastSet -LT $pwdlastChanged;
$UsersWithNeverExpiringPasswords = Get-ADUser -SearchBase $ou -filter * -properties PasswordNeverExpires | Where-Object PasswordNeverExpires -EQ $True;
$UsersWhichCannotChangePasswords = Get-ADUser -SearchBase $ou -filter * -properties CannotChangePassword | Where-Object CannotChangePassword -EQ $True;


if (!($UsersWhichMustChangePasswords -eq $null))
{
    $UsersWhichMustChangePasswords   | Select-Object Surname,GivenName,UserPrincipalName,mail,PasswordLastSet | Out-GridView -Title "Users which must change there password:" -Wait;
    $OUTPUT= [System.Windows.Forms.MessageBox]::Show("Force those Users to change there passwords?" , "Status" , 4) 
    if ($OUTPUT -eq "YES" ) 
    {
        $UsersWhichMustChangePasswords | Out-GridView -PassThru | Set-ADUser -ChangePasswordAtLogon:$True;
    }
}


if (!($UsersWithNeverExpiringPasswords -eq $null))
{
    $UsersWithNeverExpiringPasswords | Select-Object Surname,GivenName,UserPrincipalName,mail,PasswordNeverExpires | Out-GridView -Title "Users which have the PasswordNeverExpires attribute set:" -Wait;
    $OUTPUT= [System.Windows.Forms.MessageBox]::Show("Remove the PasswordNeverExpires Attribute?" , "Status" , 4) 
    if ($OUTPUT -eq "YES" ) 
    {
        $UsersWithNeverExpiringPasswords | Out-GridView -PassThru | Set-AdUser -PasswordNeverExpires:$False;
    }
}


if (!($UsersWhichCannotChangePasswords -eq $null))
{
    $UsersWhichCannotChangePasswords | Select-Object Surname,GivenName,UserPrincipalName,mail,CannotChangePassword | Out-GridView -Title "Users which have the CannotChangePassword attribute set:" -Wait;
    $OUTPUT= [System.Windows.Forms.MessageBox]::Show("Remove the CannotChangePassword attribute?" , "Status" , 4) 
    if ($OUTPUT -eq "YES" ) 
    {
        $UsersWhichCannotChangePasswords | Out-GridView -PassThru | Set-ADUser -CannotChangePassword:$False;
    }
}
