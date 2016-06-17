$scripthome = Get-Location

"
"
write-host "Coming soon to a control panel near you!"
"
"
do {
    $var = read-host -prompt "press x to go back"} until ($var -eq "x")
    if ($var -eq "x") {invoke-expression $scripthome\exchangecontrolpanel.ps1}