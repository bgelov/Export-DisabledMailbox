# Export and remove from Exchange old mailboxes
# Check script body and change it for you task

# Exchange account with export rights
$exchange_admin = 'exchange-admin@BGELOV.ru'
# Exchange server
$exchange_server = "EXCHANGE-SERVER"
# Domain name in some formats
$domain_name = "BGELOV\\"
$domain_name_simple = "BGELOV"
# Work directory
$path = 'D:\automation'
# Path for export
$export_path = '\\server\pst'
# Database name with dismiss users
$DisabledMdb_Name = "DisabledMDB"

$mailboxTrue = ""
$mailboxFalse = ""
$DisabledMDB = ''


if (!(Test-Path "$path\dismiss_users.txt")) {
	
	# Write your algorithm for getting dismissed users

    Write-Host "=== Get users who were fired more than a certain date ===" -ForegroundColor Yellow

        try
        {
			# You can using database or other way for getting dismissed users

        }
        catch [System.Data.Odbc.OdbcException]
        {
            $_.Exception
            $_.Exception.Message
            $_.Exception.ItemName
        }


    "" > "$path\dismiss_users.txt"
}


# Connect to Exchange
$UserCredential = Get-Credential $exchange_admi
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$exchange_server/PowerShell/" -Authentication Kerberos -Credential $UserCredential
Import-PSSession $Session


# Check mailboxes for dismissed users
if (!((Test-Path "$path\mailboxTrue.txt") -and (Test-Path "$path\mailboxFalse.txt"))) {

    Write-Host "=== Check mailboxes for dicmissed users ===" -ForegroundColor Yellow

    $list = gc "$path\dismiss_users.txt"
    foreach ($l in $list) {
        if (get-mailbox $l -ErrorAction SilentlyContinue) { $mailboxTrue += "$l`n" } else { $mailboxFalse += "$l`n" }  
    }

    $mailboxTrue > "$path\mailboxTrue.txt"
    $mailboxFalse > "$path\mailboxFalse.txt"
}


# Getting users in DisabledMDB
    Write-Host "=== Getting users in DisabledMDB ===" -ForegroundColor Yellow

if (Test-Path "$path\mailboxTrue.txt") {
    $mailboxTrue = gc "$path\mailboxTrue.txt"

    foreach ($mbx in $mailboxTrue) {
        if ($mbx) {
            Write-Host "$mbx"
            $nss = (get-mailbox $mbx).database
            Write-Host "$nss"

            if ("$nss" -eq 'DisabledMDB') { $DisabledMDB += "$mbx`n" }
             
        }
    }

    $DisabledMDB > "$path\DisabledMDB.txt"
}
  

# Export to .pst

Write-Host "=== Starting export to .pst ===" -ForegroundColor Yellow

$result_file = "$path\not_export.txt"
$result_file_OK = "$path\export.txt"
# OU if you need some exclude rules
$OU_contacts = 'OU=Contacts,DC=bgelov,DC=ru'
$OU_contacts2 = 'bgelov.ru/Contacts'
$OU_contacts3 = 'bgelov.ru/NewUsers'
$OU_disableU = 'bgelov.ru/Archive'
$OU_subworkers2 = 'bgelov.ru/SubUsers'

if (Test-Path $result_file) { Remove-Item -Path $result_file -Force }
if (Test-Path $result_file_OK) { Remove-Item -Path $result_file_OK -Force }

# Disabled users from DisabledMDB.txt

$user = gc "$path\$DisabledMdb_Name.txt"

$curDate = (Get-Date).AddMonths(-3)
$Date3Month = Get-Date $curDate -f "yyyy-MM-dd HH:mm:ss"


function check-sendas ($u)
{
    $login = $login_enabled = $null
    $flag = $false
    $u_mbx_SendAs = (Get-Mailbox -identity $u | Get-ADPermission | where {($_.ExtendedRights -like "*Send-As*") -and ($_.IsInherited -eq $false) -and -not (($_.User -like "NT AUTHORITY\SELF") -or ($_.User -like "$domain_name_simple\Message Processors") -or ($_.User -like "S-1-5-21-1069761793-*"))}).user
    if ($u_mbx_SendAs) {
        foreach ($u_mbx_S in $u_mbx_SendAs) {
            $login = $u_mbx_S
            $login = $login -replace $domain_name,""
            $login_enabled = (Get-ADUser $login -Properties Enabled | select Enabled).Enabled
            if ($login_enabled -eq $false) { } else { $flag = $true }
            $login = $login_enabled = $null    
        }
    }
    if ($flag -eq $true) { return $u_mbx_SendAs } else { return $null }   
}


function check-fullaccess ($u)
{
    $login = $login_enabled = $null
    $flag = $false 
    $u_mbx_perm = (Get-MailboxPermission $u | where {($_.AccessRights -eq "FullAccess") -and ($_.IsInherited -eq $False) -and -not (($_.User -like "NT AUTHORITY\SYSTEM") -or ($_.User -like "NT AUTHORITY\SELF") -or ($_.User -like "$domain_name_simple\ExchangeAdministrator") -or ($_.User -like "$domain_name_simple\Administrator") -or ($_.User -like "$domain_name_simple\Domain Admins") -or ($_.User -like "$domain_name_simple\Enterprise Admins"))}).user
    if ($u_mbx_perm) {
        foreach ($u_mbx_p in $u_mbx_perm) {
            $login = $u_mbx_p
            $login = $login -replace $domain_name,""
            $login_enabled = (Get-ADUser $login -Properties Enabled | select Enabled).Enabled
            if ($login_enabled -eq $false) { Remove-MailboxPermission -Identity $u -User $login -AccessRights FullAccess -InheritanceType All -Confirm:$false } else { $flag = $true }
            $login = $login_enabled = $null    
        }
        $u_mbx_perm = (Get-MailboxPermission $u | where {($_.AccessRights -eq "FullAccess") -and ($_.IsInherited -eq $False) -and -not (($_.User -like "NT AUTHORITY\SYSTEM") -or ($_.User -like "NT AUTHORITY\SELF") -or ($_.User -like "$domain_name_simple\ExchangeAdministrator") -or ($_.User -like "$domain_name_simple\Administrator") -or ($_.User -like "$domain_name_simple\Domain Admins") -or ($_.User -like "$domain_name_simple\Enterprise Admins"))}).user
    }
    if ($flag -eq $true) { return $u_mbx_perm } else { return $null }
}


foreach ($u in $user) {

    $u_mbx = $u_acc = $u_mbx_frwrd = $u_mbx_frwrd_DN = $u_mbx_SendAs = $u_mbx_perm = $null    

    # Exclude users
    if (!(($u -eq 'TESTuser') -or ($u -eq 'TESTuser2') -or ($u -eq 'TESTuser3') -or ($u -eq ''))) { 

        # Check user mailbox
        if ($u_mbx = get-mailbox $u -ErrorAction SilentlyContinue | select SamAccountName, ForwardingAddress) { 
        
                # Check enabled or disabled account
                $u_acc = Get-ADUser $u -Properties Enabled, extensionAttribute6 | select Enabled, PasswordLastSet, extensionAttribute6
                
                #if (!($u_acc.Enabled)) {
			    # Checkmeddage count
                $message_count = get-mailbox $u -ErrorAction SilentlyContinue | Get-MailboxStatistics | select itemcount -ExpandProperty itemcount
                $message_count

                if (!($u_acc.Enabled)) { 
				# Check users with specific attribute and less count of messages in mailbox
                #if ((!($u_acc.Enabled) -and ($u_acc.extensionAttributeX -eq 1) -and ($message_count -le 50))) {
                #if ((!($u_acc.Enabled) -and ($u_acc.extensionAttributeX -eq 1) -and ($message_count -le 50))) {

                    # Check date when user changed password
                    if (!($u_acc.PasswordLastSet -gt $Date3Month)) {

                        # Check last login
                        #if ((Get-MailboxStatistics $u | select LastLogonTime).LastLogonTime -lt $Date3Month) { 
                
                            # Check message forwarfing. If forwarding to contact or dismissed user, then continue
                            $u_mbx_frwrd = $u_mbx.ForwardingAddress
                            $u_mbx_frwrd_DN = $u_mbx_frwrd.Parent.DistinguishedName
                            if ((!($u_mbx_frwrd)) -or ($u_mbx_frwrd_DN -like "*$OU_contacts") -or ($u_mbx_frwrd -like "$OU_contacts2*") -or ($u_mbx_frwrd -like "$OU_contacts3*") -or ($u_mbx_frwrd -like "$OU_disableU*") -or ($u_mbx_frwrd -like "$OU_subworkers2*")) {

                                # Check SendAs rights
                                #$u_mbx_SendAs = check-sendas $u
                                #if (!$u_mbx_SendAs) {

                                    # Check FullAccess rights on mailbox
                                    $u_mbx_perm = check-fullaccess $u
                                    if (!$u_mbx_perm) {

                                        # Check if this mailbox already with export process...
                                        $u_status = (Get-MailboxExportRequest -Name $u | select status).status
                                        if ((!$u_status) -or (!(Test-Path "$export_path\$u.pst"))) {


                                            # Export mailbox
                                            Write-Host "Start export to $export_path\$u.pst" -ForegroundColor Green
                                            New-MailboxExportRequest -Name $u -AcceptLargeDataLoss -Confirm:$False -BadItemLimit 999 -mailbox $u -FilePath "$export_path\$u.pst"

                                            Add-Content "$u" -Path $result_file_OK


                                        } else { Write-Host "Mailbox $u already in export process" -ForegroundColor Yellow
                                                 Add-Content "$u" -Path $result_file_OK } 


                                    } else { Write-Host "On $u assigned FullAccess rights!`n$u_mbx_perm" -ForegroundColor Yellow
                                             Add-Content "$u" -Path $result_file }
                   

                                #} else { Write-Host "On $u mailbox settings up Send-As rights!`n$u_mbx_SendAs" -ForegroundColor Yellow 
                                #         Add-Content "$u" -Path $result_file }

                
                            } else { Write-Host "For $u mailbox settings up mail forwarding to $u_mbx_frwrd!" -ForegroundColor Yellow 
                                     Add-Content "$u" -Path $result_file }    
                
                
                        #} else { Write-Host "Login in this account less the 3 mounth ago!" -ForegroundColor Yellow 
                        #         Add-Content "$u" -Path $result_file }


                    } else { Write-Host "Password for $u changed less then 3 months ago!" -ForegroundColor Yellow 
                             Add-Content "$u" -Path $result_file }    
                                
            
                } else { Write-Host "Mailbox $u is in enabled state" -ForegroundColor Yellow 
                         Add-Content "$u" -Path $result_file }    
    
    
        } else { Write-Host "У $u нет почтового ящика!" -ForegroundColor Yellow 
                 $user = $user -replace $u,$null }  
    
    } # Excludes
    
}

# Delete mailboxes which doesn't exist
$user > "$path\NewDisabledMDB2.txt"
Remove-Item "$path\NewDisabledMDB.txt" -Force
Rename-Item "$path\NewDisabledMDB2.txt" "NewDisabledMDB.txt"


# Disable mailboxes what succesfully export to pst
#https://technet.microsoft.com/ru-ru/library/aa995948(v=exchg.160).aspx

$path = 'D:\automation'
$result_file_OK = "$path\export.txt"

if (Test-Path $result_file_OK) {

    Get-MailboxExportRequest | Get-MailboxExportRequestStatistics

    Write-Host "=== Disable mailboxes what succesfully export to pst ===" -ForegroundColor Yellow

    $user = gc -Path $result_file_OK

    do
    {
        $flag = $false
        foreach ($u in $user) {
    
            $u_status = $null
            $u_status = (Get-MailboxExportRequest -Name $u | select status).status

            if ($u_status) {

                # Check export process
                if ($u_status -eq 'Completed') {

                    if (Test-Path "$export_path\$u.pst") { 
            
                        Write-Host "Disable mailbox $u" -ForegroundColor Green
                        Disable-Mailbox $u -Confirm:$false
                        Get-MailboxExportRequest -Name $u | Remove-MailboxExportRequest -Confirm:$false
        
                    } else { Write-Host "No exist .pst file for $u" -ForegroundColor Yellow } 

                } elseif ($u_status -eq 'InProgress') { 
                    Write-Host "Export $u mailbox in process..." -ForegroundColor DarkYellow
                    $flag = $true
                } elseif (($u_status -eq 'Failed') -or ($u_status -eq 'FailedOther')) { Write-Host "Export error for $u mailbox" -ForegroundColor Red 
                    Get-MailboxExportRequest -Name $u | Remove-MailboxExportRequest -Confirm:$false
                } elseif ($u_status -eq 'Queued') { 
                    Write-Host "$u in queue" -ForegroundColor Yellow
                    $flag = $true
                } else { Write-Host "$u - $u_status" -ForegroundColor Yellow }

            } else { # Write-Host "No export task for $u mailbox" -ForegroundColor Red 
            
                    }
        }
        if ($flag -eq $true) { sleep -Seconds 300 }
        Write-Host "===========================================================" -ForegroundColor Yellow
    }
    while ($flag -eq $true)

    Write-Host "=== Script Ended ===" -ForegroundColor Yellow

    Get-MailboxExportRequest | Get-MailboxExportRequestStatistics

} else { Write-Host "No mailboxes for export and disabling" -ForegroundColor Green }

Remove-PSSession $Session
