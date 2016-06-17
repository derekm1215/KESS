##Contacts##
 $ErrorActionPreference = 'SilentlyContinue'
 $scripthome = Get-Location

Do {
Write-Host "
----------Contacts----------

1 = View Contacts
2 = New Contact
3 = Edit Contact - not complete
4 = Delete Contact

x = Home

---------------------------"
"
"
$choice5 = read-host -prompt "Select number & press enter"
"
"
} until ($choice5 -eq "1" -or $choice5 -eq "2" -or $choice5 -eq "3" -or $choice5 -eq "4" -or $choice5 -eq "x")

Switch ($choice5) {
    ##View Contacts## 
    "1" {
    $contact = read-host -prompt "Enter contact name here, or use wildcard
    "
    
    if ((get-mailcontact "$contact") -eq ($null)) 
        {do {$contact = read-host -prompt "That contact cannot be found, please try another"}
    until ((get-mailcontact "$contact") -ne ($null))




    get-mailcontact $contact | ft DisplayName,PrimarySmtpAddress
    "
    "
    do {
    $var = read-host -prompt "press x to go back"} until ($var -eq "x")
    if ($var -eq "x") {invoke-expression $scripthome\contacts.ps1}
    
    }}
    
    
    
    ##New Contact##
    "2"{
    
    $contact = read-host -prompt "Enter unique contact name (Spaces are allowed)"
    
    #check to ensure contact name is unique
    if ((get-mailcontact "$contact") -ne ($null)) 
        {do {$contact = read-host -prompt "That name is in use, please enter another"}
    until ((get-mailcontact "$contact") -eq ($null))}

    $email = read-host -prompt "Enter contact's email address"
    
    #check to see if user is on server, allows user to enter Q to quit to contacts menu
        if ((get-mailcontact "$email") -ne ($null)) 
        {do {$email = read-host -prompt "That email is already in the system, enter another or enter Q to go back"}
    until ((get-mailcontact "$email") -eq ($null))
        
        if ($email -eq 'Q') {invoke-expression $scripthome\contacts.ps1}}


    ""
    
    new-mailcontact -name "$contact" -externalemailaddress "$email"
    "
    "

    do {
    $var = read-host -prompt "press x to go back"} until ($var -eq "x")
    if ($var -eq "x") {invoke-expression $scripthome\contacts.ps1}}

    ##Edit Contact##
    "3" {set-mailbox
    
    do {
    $var = read-host -prompt "press x to go back"} until ($var -eq "x")
    if ($var -eq "x") {invoke-expression $scripthome\contacts.ps1}}
    
    ##Delete Contact##
    "4" {
    
    $contact = read-host -prompt "Enter contact alias"
    remove-mailcontact $contact
    "
    "
    do {
    $var = read-host -prompt "press x to go back"} until ($var -eq "x")
    if ($var -eq "x") {invoke-expression $scripthome\contacts.ps1}}
    
    ##Home##
    "x" {Invoke-Expression $scripthome\exchangecontrolpanel.ps1}
}