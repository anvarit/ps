$groupname = Read-Host "Please enter the DFS group"
$srccomputer = Read-Host "Please enter the source DFS server"
$destcomputer = Read-Host "Please enter the receiving DFS server"
Get-DfsrBacklog -GroupName $groupname -FolderName * -SourceComputerName $srccomputer -DestinationComputerName $destcomputer -Verbose