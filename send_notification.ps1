﻿<#
.SYNOPSIS
	Display pop-up message to all logged in users on remote computers
.DESCRIPTION
	This function/script uses MSG.exe to send a message to all logged in users of a single, or multiple, computers
.PARAMETER ComputerName
	Name of desired computer/s
.PARAMETER Message
	Message to be sent/displayed
.PARAMETER TimeInSeconds
	The number of seconds the message will be displayed before it disappears
.INPUTS
	System.String
.OUTPUTS
	N/A
.EXAMPLE
	.\Send-MessagePopUp.ps1 -Computer SERVER01 -Message 'All users will be logged off in 5 minutes for system maintenance' -Verbose

#>

[CmdletBinding()]
param (
	[Parameter(Position = 0,
			   Mandatory = $true)]
	[System.String[]]$ComputerName,

	[Parameter(Position = 1,
			   Mandatory = $true)]
	[System.String]$Message,

	[Parameter(Position = 2,
			   Mandatory = $false)]
	[System.String]$TimeInSeconds
)

BEGIN {

	# Setting global EA preference since we are calling a system binary and not a PowerShell script, function or cmdlet that support -ErrorAction
	$ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop

} # end BEGIN block

PROCESS {

	foreach ($computer in $ComputerName) {
		$i++

		if (Test-Connection -ComputerName $computer -Count 2 -Quiet) {
			try {

				Write-Verbose -Message "Sending message to $computer"
				if ($TimeInSeconds) {

					MSG.exe * /SERVER:$computer /TIME:$TimeInSeconds $Message

				} else {

					MSG.exe * /SERVER:$computer $Message

				} # end if/else

			} catch {

				Write-Warning -Message "$computer - $_"

			} # end try/catch

		} else {

			Write-Warning -Message "[$computer] - Unreachable via Ping"

		} # if/else

		# Write total progress to progress bar
		$totalComputers = $ComputerName.Length
		$percentComplete = [int](($i / $totalComputers) * 100)
		Write-Progress -Activity "Working..." -CurrentOperation "$percentComplete% Complete" -Status "Percent Complete" -PercentComplete $percentComplete

	} # foreach $computer

} # end PROCESS block

END {

} # end END block