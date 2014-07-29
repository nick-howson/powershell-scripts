# Gets information for a computer listed in the AD

# Import the AD module
Import-Module ActiveDirectory
Do {
# get Name
$hostname = Read-host "Enter a hostname"
Get-ADComputer $hostname

# little menu
$menuTitle 		= ""
$menuMessage	= "Do you want to query another hostname?"

$menuOptQuery	= New-Object System.Management.Automation.Host.ChoiceDescription "&Query",`
"Runs another query"
$menuOptExit	= New-Object System.Management.Automation.Host.ChoiceDescription "E&xit",`
"Exits"

# As options are added the return code follows the sequence in which they are added in this case QUery is code 0 and exit is code 1
# The first entered option is the default choice
$menuOptions	= [System.Management.Automation.Host.ChoiceDescription[]]($menuOptQuery, $menuOptExit)

$result = $host.ui.PromptForChoice($menuTitle, $menuMessage, $menuOptions, 0) 
}
Until ($result -eq "1")