# Export-DisabledMailbox
Advanced script for export old mailboxes from Microsoft Exchange Server to .pst files

## Some usefull commands
```
# Free space in databases
Get-MailboxDatabase -status |ft server,name, AvailableNewMailboxSpace -AutoSize

# Remove disabled mailboxes artifacts for DisabledMDB database
Get-MailboxStatistics -Database DisabledMDB | where {$_.DisconnectReason -eq "disabled"} | foreach {Remove-StoreMailbox -Database $_.database -Identity $_.mailboxguid -MailboxState disabled}

Get-MailboxStatistics -Database DisabledMDB | where {$_.DisconnectReason -eq "Softdeleted"} | foreach {Remove-StoreMailbox -Database $_.database -Identity $_.mailboxguid -MailboxState disabled}

Get-MailboxDatabase DisabledMDB | foreach {Get-MailboxStatistics -Database $_.identity} | ForEach { Update-StoreMailboxState -Database $_.Database -Identity $_.MailboxGuid -Confirm:$false}
```
