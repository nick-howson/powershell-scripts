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
# Import the AD module
Import-Module ActiveDirectory
$output_file_name		= "staleComputers"
# More information on date and time fuctions see - http://technet.microsoft.com/en-us/library/ff730960.aspx
$dateCurrent = Get-Date
$date6 = $dateCurrent.AddMonths(-6)
$date12 = $dateCurrent.AddMonths(-12)
$date18 = $dateCurrent.AddMonths(-18)

################################
# Generate CSV reports section #
################################
# Get all AD computers with lastLogonTimestamp of 6 months or older
Get-ADComputer -Filter {LastLogonTimeStamp -lt $date6} -Properties LastLogonTimeStamp |
# Output hostname and lastLogonTimestamp into CSV
select-object Name,@{Name="Stamp"; Expression={[DateTime]::FromFileTime($_.lastLogonTimestamp)}} | export-csv $output_file_name"-6months.csv" -notypeinformation
# Get all AD computers with lastLogonTimestamp of 6 months or older
Get-ADComputer -Filter {LastLogonTimeStamp -lt $date12} -Properties LastLogonTimeStamp |
# Output hostname and lastLogonTimestamp into CSV
select-object Name,@{Name="Stamp"; Expression={[DateTime]::FromFileTime($_.lastLogonTimestamp)}} | export-csv $output_file_name"-12months.csv" -notypeinformation
# Get all AD computers with lastLogonTimestamp of 6 months or older
Get-ADComputer -Filter {LastLogonTimeStamp -lt $date18} -Properties LastLogonTimeStamp |
# Output hostname and lastLogonTimestamp into CSV
select-object Name,@{Name="Stamp"; Expression={[DateTime]::FromFileTime($_.lastLogonTimestamp)}} | export-csv $output_file_name"-18months.csv" -notypeinformation
# press any key to continue prompt mimicks the old pause from batch files
Write-Host "Press any key to continue ..."
$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")