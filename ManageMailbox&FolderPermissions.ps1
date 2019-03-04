<# Script to Get/Remove mailbox and its folder permissions

*DESCRIPTION:
This simple PS script helps to check and remove permissions of a user on another mailbox or each of its folders.
You can either choose to 'Get' information of the folders that were granted with any access rights or choose to remove the same access rights (on confirmation).
Or choose the other option to check 'Get' and remove(on confirmation) mailbox permissions (delegations) of one user grated over another user mailbox.

Warning: Once confirmed on first prompt while removing permissions, all the selected accesses will be revoked without any further notice on confirmation.

*Follow the INSTRUCTIONS below to perform the task - AYOR...
Extract the '.zip' file to a folder
Open Exchange Management Shell and navigate to the path of this script
For example: 
------------
PS C:\Users\KrishKT\Desktop\ScriptsTest>.\ManageMbxFolderPermissions.ps1
Hit 'Enter' to run the script
Follow on-screen prompts!

* @KrishKT | #XploreMore 
#>

#exec
function Search-Stats
{Write-Host "Do you want to proceed with checking/removing of mailbox folder permissions?"
$fmanage = Read-Host -Prompt "Enter 1 to 'Check' or 2 to 'Remove' all folder permissions of $usr on $mbx mailbox folders"
switch ($fmanage)
{{(($fmanage -eq "1") -or ($fmanage -eq "2"))}
{$mbxfdrs=(Get-MailboxFolderStatistics $mbx).folderpath
$count=$mbxfdrs.count
Write-host "Total number of mailbox folders (including subfolders)of $mbx : $count"
Write-host "Checking $usr's permission on all folders of $mbx..."
foreach ($fdrs in $mbxfdrs)
{$folder = $mbxfdrs -replace("/","\")}
$identity = $folder| foreach {"$($mbx):$_"}
$permissions = $identity | foreach{Get-MailboxFolderPermission "$_" -User $usr -ErrorAction silentlycontinue}
if ($fmanage -eq "1" -and ($permissions))
{
Write-Host -f "yellow" "$usr has permissions on the following mailbox folders of $mbx :"
$permissions | ft -autosize -wrap}
elseif ($fmanage -eq "2" -and ($permissions))
{
Write-Host -f "red" -b white "You are about to REMOVE ALL existing permissions of $usr from ALL mailbox folders of $mbx ?"
$confirm = Read-Host "Require Confirmation [Y(yes)/N(no)]"
if ($confirm -eq "yes" -or $confirm -eq "y")
{$permissionr = $identity | foreach{Remove-MailboxFolderPermission "$_" -User $usr -Confirm:$false -ErrorAction silentlycontinue}
Write-Host -f "yellow" "Existing permissions of user: $usr were removed from the following mailbox folders of $mbx :"
$permissions | ft -autosize -wrap
Write-Host -f "green" "No more permissions exist now for $usr on $mbx's mailbox folders!"} 
elseif ($confirm -eq "no" -or $confirm -eq "n")
{Write-Host "Task Aborted!"}
elseif ($confirm -ne "no" -or $confirm -ne "n" -or $confirm -ne "yes" -or $confirm -ne "y")
{
Write-Host "False input :-( Retry!"}}
elseif ((($fmanage -eq "1") -and (!$permissions)) -or (($fmanage -eq "2") -and (!$permissions))) 
{Write-Host -f "yellow" "There is no existing permission entry found for user:$usr on $mbx's mailbox folders"}}
default {"False input :-( Retry!"}}
}
function Search-Delegate
{$mbxaccess = Read-Host "Enter 1 to 'Check' and 2 to 'Remove' selected or all $usr's permission(s) on $mbx's mailbox"
if ($mbxaccess -eq "1" -or $mbxaccess -eq "2")
{
Write-Host "Checking for ALL existing mailbox permissions of $usr on $mbx mailbox..."
$dnu = (Get-Mailbox $usr).distinguishedname
$maccess = Get-MailboxPermission -Identity $mbx -User $usr
$SendAs = Get-Mailbox $mbx | Get-ADPermission -User $usr | ?{$_.ExtendedRights -like 'Send*'} | select user, identity, deny, accessrights
$SendOn = Get-Mailbox $mbx | Select -ExpandProperty GrantSendOnBehalfTo | ?{$_.DistinguishedName -eq $dnu} | select name,parent
$idusr = (Get-Mailbox $usr).name
$idS = $SendOn.name
$nsend = $idS -contains $idusr
if ((!$maccess.accessrights) -and ($nSend -ne "True") -and (!$SendAs))
{
$del = "nope"
Write-Host -f "yellow" "None of the mailbox permission exist for $usr on $mbx's mailbox"}
if (($mbxaccess -eq "1") -and ($del -ne "nope"))
{Write-Host -f red "NOTE: Only granted permissions of $usr on $mbx's mailbox will be retrived with details, BELOW respective access rights title."
Write-Host -f blue "FullAccess:" 
$maccess | fl
Write-Host -f blue "Send-As:" 
$SendAs | fl
Write-Host -f blue "SendOnBehalf:"
switch ($nsend)
{{$nSend -eq "True"} {$SendOn | fl}}}
elseif (($mbxaccess -eq "2") -and ($del -ne "nope"))
{$afs = @("a", "f", "s")
$cafs = @("y", "yes", "n", "no")
Write-Host " "
Write-Host "`n'A' - to remove ALL granted mailbox permissions of $usr on $mbx mailbox"
Write-Host "'F' - to remove $usr's FullAccess permission on $mbx mailbox"
Write-Host "'S' - to $usr's Send permission on $mbx mailbox"
$rmaccess = Read-Host "Choose one of the above options[A/F/S] to proceed"
Write-Host -f red -b white "You are about to REMOVE the above selected mailbox permission(s) of $usr from $mbx"
$rconfirm = Read-Host "Require Confirmation [Y(yes)/N(no)]"
switch ($rmaccess -and $rconfirm)
{{($afs -notcontains $rmaccess) -or ($cafs -notcontains $rconfirm)} {"False input :-( Retry!"}
{($rconfirm -eq "n" -or $rconfirm -eq "no")} {"Task Aborted!"}
{($afs -contains $rmaccess) -and ($cafs -contains $rconfirm)}
{if (!$maccess.accessrights -and ($rconfirm -ne "n" -and $rconfirm -ne "no") -and ($rmaccess -ne "s"))
{Write-Host -f yellow "No FullAccess permission exist for user: $usr on $mbx's mailbox."}
elseif (($rmaccess -eq "a" -or $rmaccess -eq "f") -and ($rconfirm -eq "y" -or $rconfirm -eq "yes") -and ($maccess.accessrights))
{
Remove-MailboxPermission -Identity $mbx -User $usr -AccessRights FullAccess -InheritanceType All -Confirm:$false
Write-Host -f green "$usr's FullAccess permission on $mbx mailbox has been removed."}
switch ($rmaccess)
{{((!$SendAs) -and ($rconfirm -ne "n" -and $rconfirm -ne "no") -and ($rmaccess -ne "f"))} {Write-Host -f yellow "No Send-As permission exist for user: $usr on $mbx's mailbox."}
{(($nSend -ne "True") -and ($rconfirm -ne "n" -and $rconfirm -ne "no") -and ($rmaccess -ne "f"))} {Write-Host -f yellow "No SendOnBehalf permission exist for user: $usr on $mbx's mailbox."}
{(($rmaccess -eq "a" -or $rmaccess -eq "s") -and ($rconfirm -eq "y" -or $rconfirm -eq "yes") -and (($SendAs) -or ($nSend -eq "True")))}
{$dnm = (Get-Mailbox $mbx).distinguishedname
Remove-ADPermission -Identity "$dnm" -User $usr -ExtendedRights "Send As" -Confirm:$false
Set-Mailbox $mbx -GrantSendOnBehalfTo @{remove=$usr} -Confirm:$false
Write-Host -f "green" "$usr's Send permission on $mbx mailbox has been removed."
}}}}}
Clear-Variable del -ErrorAction SilentlyContinue}
else {Write-Host "False input :-( Retry!"}
}
function proceed
{Write-Host "`nGreetings :-) $env:Username!"
Write-Host "=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-="
Write-Host "Get/Remove mailbox and its folder permissions"
Write-Host "=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-="
$mbx = Read-Host -Prompt "Enter the alias/email address of the user whose mailbox or its folders permissions have to be checked"
$usr = Read-Host -Prompt "Enter the alias/email address of the user whose permissions on $mbx's mailbox or its folders have to be retrieved"
Write-Warning -Message "Recommended to ensure the above users exist for correct results!"
Write-Host "Do you want to proceed with managing mailbox permissions or mailbox folder permissions?"
$manage = Read-Host -Prompt "Enter 'M' to proceed with mailbox permissions or 'F' to proceed with mailbox folder permissions"
switch ($manage)
{{$manage -eq "F"} {Search-Stats}
{$manage -eq "M"} {Search-Delegate}
default {Write-Host "False input :-( Retry!"}}}
proceed
#globals
#version1.0