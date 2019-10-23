Function YesNo
{
    Param($provide)
    $a = new-object -comobject wscript.shell 
    $popup = "Do you want to Select the servers to scan ?"
    $intAnswer = $a.popup($popup, ` 
    0,$provide,4) 
        If ($intAnswer -eq 6) { 
            $answer = "yes" 
        } else { 
            $answer = $null 
        } 
    return $Answer
}
Function DisplayInfo
{
    Param($Servers)
    cls
    write-host
    $count = $servers.count
    write-host $count -ForegroundColor Yellow -NoNewline
    write-host " servers selected to investigate." -ForegroundColor Green
    write-host "The total of servers in the AD is :" -ForegroundColor Green -NoNewline
    Write-Host (get-adcomputer -filter {Operatingsystem -like "Windows Server*"} | select name).count -ForegroundColor Red
    Write-Host 
}
Function Selection
{
    #$servers = (get-adcomputer -filter {Operatingsystem -like "Windows Server*"} | sort name | select name).name
    $servers = (get-adcomputer -LDAPFilter "(&(objectCategory=computer)(operatingSystem=Windows Server*)(!serviceprincipalname=*MSClusterVirtualServer*)(!(userAccountControl:1.2.840.113556.1.4.803:=2)))" -Property name | sort-object Name).name
    if (yesNo -eq "Yes")
    {
        $servers = $servers | out-gridview -passthru
    }
    return $servers
}
 
Function DoIt
{
    Param($Servers)
    $completelist = @()
    foreach ($server in $servers) 
    {
        $serverdisp = $server.PadRight(($servers | Measure-Object -Maximum -Property Length).Maximum)
        $Skip = 0
        Write-host "Checking server " -ForegroundColor Green -NoNewline; write-host $Server -ForegroundColor Yellow -NoNewline
        if (Test-Connection $server -count 1 -Quiet)
        {
            Write-Host "         ------ (online)" -ForegroundColor Green
            $sessionId = $null
            try 
            {
                #$ErrorActionPreference = "SilentlyContinue"
                $SessionIds = quser /server:$server 2>&1
                #$ErrorActionPreference = "Continue"
            }
            catch
            { 
                Write-Host "              - " -ForegroundColor Green -NoNewline
                write-host "$serverdisp seems inaccessible." -ForegroundColor Red -NoNewline
                Write-Host " No data available." -ForegroundColor Yellow
                write-host " ---------------------- " -ForegroundColor green
                $skip = 1
            }
            if ($Skip -eq 0)
            {
                $nr = $sessionIds.count-2
                $c=0
                do
                {
                    $c++
                    $user = (($sessionIds[$c]) -split ' +')[1]
                    $session = (($sessionIds[$c]) -split ' +')[2]
                    $State = (($sessionIds[$c]) -split ' +')[3]
                    $IdleT = (($sessionIds[$c]) -split ' +')[4]
                    $LogonT = (($sessionIds[$c]) -split ' +')[5] + " " + (($sessionIds[$c]) -split ' +')[6]
                    Write-Host "              - " -ForegroundColor Green -NoNewline
                    write-host $serverdisp -ForegroundColor Yellow -NoNewline
                    Write-Host " - " -ForegroundColor Green -NoNewline
                    write-host $user -ForegroundColor Magenta -NoNewline
                    Write-Host " - " -ForegroundColor Green -NoNewline
                    write-host $session -ForegroundColor Cyan -NoNewline
                    Write-Host " - " -ForegroundColor Green -NoNewline
                    write-host $State -ForegroundColor Cyan -NoNewline
                    Write-Host " - " -ForegroundColor Green -NoNewline
                    write-host $LogonT -ForegroundColor Cyan -NoNewline
                    Write-Host " => " -ForegroundColor Green -NoNewline
                    write-host $IdleT -ForegroundColor Cyan
                    # $Ser+=@($server); $use+=@($user);$Ses+=@($session);$sta+=@($state);$Idl+=@($IdleT);$Log+=@($LogonT)
                    $completelist += @{server=$server;user=$user;session=$session;State=$state;IdleTime=$idleT;LogonTime=$logonT}
                }
                while ($c -le $nr)
                write-host " ---------------------- " -ForegroundColor green
            }
        }
        else
        {
                Write-Host "         ------ (offline)" -ForegroundColor Red
                Write-Host "              - No data available." -ForegroundColor Yellow
                write-host " ---------------------- " -ForegroundColor green
        }
    }
    <# $completeList = New-Object PSObject
    $completeList | Add-Member NoteProperty Server   $Ser
    $completeList | Add-Member NoteProperty User     $use
    $completeList | Add-Member NoteProperty Session  $ses
    $completeList | Add-Member NoteProperty State    $sta
    $completeList | Add-Member NoteProperty IdleTime $Idl
    $completeList | Add-Member NoteProperty Logon    $Log
    $completelist = @(server=$ser;user=$use;session=$ses;State=$sta;IdleTime=$idl;LogonTime=$log) #>
    return $completeList
}
Function YesNo2
{
    Param($Completelist)
    $more = @()
    $List = $completelist |% {New-Object psobject -Property $_}
    $Users = ($completelist |% {New-Object psobject -Property $_}).user | sort -Unique
    $a = new-object -comobject wscript.shell 
    $intAnswer = $a.popup("Do you want to Log Off (A) certain User(s)?", ` 
    0,"Logoff Users",4) 
    If ($intAnswer -eq 6) 
    { 
        $answer2 = $Users | Out-GridView -PassThru
        foreach ($sub in $answer2)
        {
            $more += $list | where {$_.user -eq $sub}
        }
        $answer2 = $more | select server, user, State, logontime, idletime, session | Out-GridView -PassThru
        $step = 0
 
        foreach ($log in $answer2)
            {
                write-host $log.user -ForegroundColor green -NoNewline
                write-host " logged on to device " -ForegroundColor Yellow -NoNewline
                write-host $log.server -ForegroundColor Green -NoNewline
                Write-Host " since " -ForegroundColor Yellow -NoNewline
                write-host $log.Logontime -ForegroundColor Cyan
                write-host "   °Attempt log off:" -ForegroundColor Yellow -NoNewline
                $connect = $log.server
                try{Invoke-RDUserLogoff -HostServer $connect -UnifiedSessionId $Log.session -Force ;$logof="OK"}
                catch{$logof="NOK"}
                if($Logof -eq "OK"){ write-host " log off succesful !" -ForegroundColor Green}
                else { write-host " log off failed! Please check manually." -ForegroundColor red}
 
            }
    }
}
 
 
 
# SO now let's play!
$Servers = Selection
Displayinfo $servers
$completelist = DoIt $servers
yesno2 $completelist
 