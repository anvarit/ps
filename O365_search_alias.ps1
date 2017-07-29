$alias = Read-Host "Please enter the alias, add ' before and after"

$alias = 'F3H@barclab.com'
Get-Mailbox -Identity * | Where-Object {$_.EmailAddresses -like $alias} | Format-List Identity