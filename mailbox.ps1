$erroractionpreference = 'SilentlyContinue'
$scripthome = Get-Location

##Mailbox Menu##
Do {
Write-Host "
----------Mailboxes------------------

1 = View Mailboxes
2 = Show Mailboxes with forwarders
3 = New User
4 = Edit User
5 = Unlock Locked Account

x = Home

------------------------------------"
"
"
$choice2 = read-host -prompt "Select number & press enter"
} until ($choice2 -eq "1" -or $choice2 -eq "2" -or $choice2 -eq "3" -or $choice2 -eq "4" -or $choice2 -eq "5" -or $choice2 -eq "6" -or $choice2 -eq "x")
""
   Switch ($choice2) {
    
    
    
    ##View Mailboxes## 
    "1" {$name = read-host -prompt "Enter Name or use * for wildcard"
"
"  #get the maibox for the name, if the first attempt finds nothing, you can continue trying or enter Q to break the script
    
    if ((get-mailbox $name) -eq $null) {do {$name = read-host "A mailbox with that name could not be found, please try again or enter Q to quit."}
    until (((get-mailbox $name) -ne $null) -or ($name -eq "Q"))}
""
    if ($name -ne "Q"){Get-mailbox $name}
"
"
do {
    $var = read-host -prompt "press x to go back"} until ($var -eq "x")
    if ($var -eq "x") {invoke-expression $scripthome\mailbox.ps1}}
    
    ##Show Mailboxes with Forwarders##
    "2" {
    if ((get-mailbox | select forwardingaddress | where {$_.ForwardingAddress -ne $null}).count -ne $null) {
    
    Get-mailbox | select DisplayName,ForwardingAddress | where {$_.ForwardingAddress -ne $Null} }
    

    else {write-host "There are no forwarders setup on this server"}

    do {
    $var = read-host -prompt "press x to go back"} until ($var -eq "x")
    if ($var -eq "x") {invoke-expression $scripthome\mailbox.ps1}}

    
    
    
    ##new Mailboxes##
    "3" 
   {
        $Domain = $env:userdnsdomain
    $First = read-host -prompt "First name"
    $Last = read-host -prompt "Last name"
    $name = $first + " " + $Last
    $upn = read-host -prompt "Enter email address"
     if ((get-mailbox $upn) -ne ($null)) {do {write-host "This email address is already in use, please try another."
    $upn = read-host -prompt "Enter email address"}

    until ((get-mailbox $upn) -eq ($null))}
    
    
    $password = read-host -prompt "Enter Password" -AsSecureString
    "
    "
    

    write-host "Choose an OU"
    
    "
    "
    function Select-TextItem 
{ 
PARAM  
( 
    [Parameter(Mandatory=$true)] 
    $options, 
    $displayProperty 
) 
 
    [int]$optionPrefix = 1 
    # Create menu list 
    foreach ($option in $options) 
    { 
        if ($displayProperty -eq $null) 
        { 
            Write-Host ("{0,3}: {1}" -f $optionPrefix,$option) 
        } 
        else 
       { 
          Write-Host ("{0,3}: {1}" -f $optionPrefix,$option.$displayProperty) 
       } 
        $optionPrefix++ 
    } 
    Write-Host ("{0,3}: {1}" -f 0,"To cancel")  
    [int]$response = Read-Host "Enter Selection" 
    $val = $null 
    if ($response -gt 0 -and $response -le $options.Count) 
    { 
        $val = $options[$response-1] 
    } 
    return $val 
}    
 
$values = get-organizationalunit 
$val = Select-TextItem $values "canonicalname" 

 
    
     
        
    ""
    
    ""

    #If there are less than 2 mailbox dbs, no need to select from a list
    
    If ((get-mailboxdatabase | group-object).count -lt 2)

    {new-mailbox -name $name -userprincipalname $upn -Password $password -OrganizationalUnit "$val" | out-null
    #get-mailbox -identity $upn | ft displayname,userprincipalname,organizationalunit,database
    }

    else #select from a list of Mailbox DBs

    {write-host "Which Mailbox database?"
    
    ""

        function Select-TextItem 
{ 
PARAM  
( 
    [Parameter(Mandatory=$true)] 
    $options, 
    $displayProperty 
) 
 
    [int]$optionPrefix = 1 
    # Create menu list 
    foreach ($option in $options) 
    { 
        if ($displayProperty -eq $null) 
        { 
            Write-Host ("{0,3}: {1}" -f $optionPrefix,$option) 
        } 
        else 
       { 
          Write-Host ("{0,3}: {1}" -f $optionPrefix,$option.$displayProperty) 
       } 
        $optionPrefix++ 
    } 
    Write-Host ("{0,3}: {1}" -f 0,"To cancel")  
    [int]$response = Read-Host "Enter Selection" 
    $val = $null 
    if ($response -gt 0 -and $response -le $options.Count) 
    { 
        $val = $options[$response-1] 
    } 
    return $val 
}    
 
$values = get-mailboxdatabase 
$mbdb = Select-TextItem $values "name" 
#$mbdb


new-mailbox -name $name -userprincipalname $upn -Database "$mbdb" -Password $password -OrganizationalUnit "$val" | out-null
    #get-mailbox -identity $upn | ft displayname,userprincipalname,organizationalunit,database
    }
   
   If ((get-mailbox -identity $upn) -eq $null) 
   
   {do {$password = read-host "That didn't work, I bet your password was weak, let's try again" -AsSecureString
   new-mailbox -name $name -userprincipalname $upn -Database "$mbdb" -Password $password -OrganizationalUnit "$val" | out-null
    }
    
    until ((get-mailbox -identity $upn) -ne $null)}
   
   get-mailbox -identity $upn | ft displayname,userprincipalname,organizationalunit,database
    
    "
    "
    do {
    $var = read-host -prompt "press x to go back"} until ($var -eq "x")
    if ($var -eq "x") {invoke-expression $scripthome\mailbox.ps1}}
    
  
   
   
   ##edit Mailboxes##
    "4" {Invoke-Expression $scripthome\editmailbox.ps1}

    
    
    
    <##Disable Mailboxes --- shelved since I really don't want people having this option this easy
    "5" {
    $n = read-host -prompt "Enter alias of user to be disabled"

if ((get-mailbox $n) -eq ($null)) {do {write-host "This user was not found"
        $n = read-host -prompt "Enter Email"}
        until ((get-mailbox $n) -ne ($null))}
    "
    "
    if (get-aduser $n | where enabled -eq $false){'This account has already been disabled'}
    "
    "
    disable-adaccount -identity "$n"
    "
    "
    

    get-aduser $n | fl Name,Enabled,UserPrincipalName
    "
    "
    do {
    $var = read-host -prompt "press x to go back"} until ($var -eq "x")
    if ($var -eq "x") {invoke-expression $scripthome\mailbox.ps1}}#>




  ##Unlock Account##
    "5" {
    $n = read-host -prompt "Enter alias of user to be Unlocked"

if ((search-adaccount -lockedout | where SamAccountName -eq $n) -eq ($null)) {do {write-host "A locked account by this username cannot be found"
        $n = read-host -prompt "Enter Alias"}
        until ((search-adaccount -lockedout | where SamAccountName -eq $n) -ne ($null))}
    "
    "
    unlock-adaccount -identity "$n"
    "
    "
    "Account Successfully Unlocked"
    "
    "
    do {
    $var = read-host -prompt "press x to go back"} until ($var -eq "x")
    if ($var -eq "x") {invoke-expression $scripthome\mailbox.ps1}}  
    
    
    
    ##Home##
    "x" {Invoke-Expression $scripthome\exchangecontrolpanel.ps1}
   }
    
    