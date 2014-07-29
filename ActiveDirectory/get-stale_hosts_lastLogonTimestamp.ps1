# Gets time stamps for all computers in the domain that have NOT logged in since after specified date

# Import the AD module
Import-Module ActiveDirectory

# setup
$time 					= Read-host "Enter a date in format mm/dd/yyyy"
$time 					= get-date ($time)
$date 					= get-date ($time) -UFormat %d.%m.%y
$output_file_name		= "staleComputers.csv"

# Get all AD computers with lastLogonTimestamp less than our time
Get-ADComputer -Filter {LastLogonTimeStamp -lt $time} -Properties LastLogonTimeStamp |
 
# Output hostname and lastLogonTimestamp into CSV
select-object Name,@{Name="Stamp"; Expression={[DateTime]::FromFileTime($_.lastLogonTimestamp)}} | export-csv $output_file_name -notypeinformation

# press any key to continue prompt mimicks the old pause from batch files
Write-Host "Press any key to continue ..."
$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")