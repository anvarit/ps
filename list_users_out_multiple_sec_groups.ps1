$Grps = Get-ADGroup -Filter{name -like "BarcSafe*"}

$results=Foreach ($Grp in $Grps){
  Get-ADPrincipalGroupMembership -Identity $Grp.SamAccountName |
   Get-ADGroupMember
}
$results | Export-csv 'c:\temp\secgrp.csv' -notypeinformation