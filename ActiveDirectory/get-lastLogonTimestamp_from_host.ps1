# Gets host and lastLogonTimestamp in UTC of specified host
 
# Import the AD module
Import-Module ActiveDirectory
 
# get Name
$hostname = Read-host "Enter a hostname"
 
# grab the lastLogonTimestamp attribute
Get-ADComputer $hostname -Properties lastlogontimestamp |
 
# output hostname and timestamp in human readable format
Select-Object Name,@{Name="Stamp"; Expression={[DateTime]::FromFileTime($_.lastLogonTimestamp)}}

# press any key to continue prompt mimicks the old pause from batch files
Write-Host "Press any key to continue ..."
$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")