# Gets time stamps for all computers in the domain that have NOT logged in since after specified date
 
$time 					= Read-host "Enter a date in format mm/dd/yyyy"
$time 					= get-date ($time)
$date 					= get-date ($time) -UFormat %d.%m.%y
$output_file_name		= "staleComputers.csv"

# Get all AD computers with lastLogonTimestamp less than our time
Get-ADComputer -Filter {LastLogonTimeStamp -lt $time} -Properties LastLogonTimeStamp |
 
# Output hostname and lastLogonTimestamp into CSV
select-object Name,@{Name="Stamp"; Expression={[DateTime]::FromFileTime($_.lastLogonTimestamp)}} | export-csv $output_file_name -notypeinformation
