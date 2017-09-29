$computer = Read-Host "Please enter the receiving DFS server"
get-dfsrstate -ComputerName $computer | sort UpdateState -Descending | ft path,inbound,updatestate,sourcecomputername -auto -wrap