# Wayne T Taylor, March-2019
# For Office 365 
# run inside the microsoft exchnage online powershell module
# You Need to connect to 365 with the commands below for this to work.

# Will get a list of all accounts, then check if they have a mailbox and 
# of so they are then chcked to see if they have any rules and outputted 
# to the screen with the name of the mailbox above the rules list.

Install-Module -Name CredentialManager

$creds = Get-StoredCredential

Connect-EXOPSSession -Credential $creds
connect-msolservice -Credential $creds

# Get all users login names and store in $users for later use
$users = Get-MsolUser | Select-Object -Property UserPrincipalName

# Loop the users and check the inbox rules.
foreach($user in $users)
{
    # Check if a the account has a mailbox 
    $exist = [bool](Get-mailbox $user.UserPrincipalName -erroraction SilentlyContinue);
    
    # If account has mailbox check to see if it has any rules applied
    if ($exist -eq "True")
    {
        # Output username (so you know which mailbox has the rules)
        Write-Host $user.UserPrincipalName
        Write-Host "-------------------------------------------------"

        # Query Mailbox for any rules and outout
        get-InboxRule -Mailbox $user.UserPrincipalName
    }
}