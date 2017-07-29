

$server = "192.168.1.152"
$user = "user"
$pwd = "password"

Connect-VIServer $server -User $user -Password $pwd

$vms = Get-VM
write "`nVM's with CD-ROM device type set to 'Datastore ISO file' :"
foreach ($vm in $vms | where { $_ | Get-CDDrive | where { $_.ISOPath -like "*.ISO*"} | Set-CDDrive -Nomedia -confirm:$false })  {
	write $vm.name
	
}
