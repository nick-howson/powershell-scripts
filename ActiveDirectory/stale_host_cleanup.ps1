<#
.SYNOPSIS
	Gets a list of all the stale computers (Not seen in 6+ months) timestamps them in the Description field and moves them to the appropriate AD Objects
.DESCRIPTION
	Uses the AD modle to get access to the Get-ADComputer & Move-ADObject cmdlets to allow the user to query a computer for its last seen date, then move it to an OU depending on its age. 
	
	This script is designed to be run from the task shedualer as a weekly task.
.NOTES
Author: Nick Howson <nick.howson@gmail.com>
#>

#Setup
# More information on date and time fuctions see - http://technet.microsoft.com/en-us/library/ff730960.aspx
$dateCurrent = Get-Date

<# quick string date to DateTime converter
$string = "19/03/2010 14:11"
$dateformat = "%d/MM/yyyy %H:mm"
[datetime]::ParseExact($string,$dateformat,$null)
#>
