##Resources##

$ErrorActionPreference = 'SilentlyContinue'
$scripthome = Get-Location

Do {
Write-Host "
----------Resources-------------------

1 = View Resources
2 = New Resource
3 = Common Resource Calendar Settings
4 = Disable Resource

x = Home

--------------------------------------"
"
"
$choice4 = read-host -prompt "Select number & press enter"
"
"
} until ($choice4 -eq "1" -or $choice4 -eq "2" -or $choice4 -eq "3" -or $choice4 -eq "4" -or $choice4 -eq "x")

Switch ($choice4) {
    
    ##View Resources## 
    "1" {
    
    $room = (read-host "Enter room name or use * for wildcard")
    get-mailbox "$room" -RecipientTypeDetails RoomMailbox | ft name,alias
    
    "
    "
    do {
    $var = read-host -prompt "press x to go back"} until ($var -eq "x")
    if ($var -eq "x") {invoke-expression $scripthome\resources.ps1}
    }
    
    ##new Resource 
    "2" {
    
    $alias = read-host -prompt "Enter unique resource alias (no spaces)"
    
    #check to ensure resource is unique
    if ((get-mailbox "$alias") -ne ($null)) 
        {do {$alias = read-host -prompt "That name is in use, please enter another"}
    until ((get-mailbox "$alias") -eq ($null))}

    $roomdn = read-host -prompt "Enter room display name"
    $book = read-host -prompt "Enter days in advance room can be booked (Default is 180)"
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
$ou = Select-TextItem $values "canonicalname"

 
    
     
        
    ""
    
    ""

    #If there are less than 2 mailbox dbs, no need to select from a list
    
    If ((get-mailboxdatabase | group-object).count -lt 2)

    {New-Mailbox -Name $alias -DisplayName "$roomdn" -organizationalunit "$OU" -room | out-null
    get-mailbox -name $alias}

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


New-Mailbox -Name $alias -DisplayName "$roomdn" -database "$mbdb" -organizationalunit "$OU" -room | out-null
    get-mailbox "$roomdn"}
    
    "
    "
    do {
    $var = read-host -prompt "press x to go back"} until ($var -eq "x")
    if ($var -eq "x") {invoke-expression $scripthome\resources.ps1}
    }

##Custom Resource Calendar Settings##
    "3" {
    
    $alias = read-host -prompt "Resource to edit"

    #check to see if user is on server
        if ((get-mailbox $alias) -eq ($null)) {do {write-host "This room was not found"
        $alias = read-host -prompt "Enter Room Alias"}
        until ((get-mailbox $alias) -ne ($null))}

    
    
    
    
    Set-calendarprocessing -identity $alias -automateprocessing:autoaccept -deleteattachments $false -deletesubject $false -addorganizertosubject $false -deletecomments $false
    
    $alias = $alias + ':'
    Set-MailboxFolderPermission -identity $alias\Calendar -User Default -AccessRights Reviewer
    "
    "
    do {
    $var = read-host -prompt "press x to go back"} until ($var -eq "x")
    if ($var -eq "x") {invoke-expression $scripthome\resources.ps1}
    }

    
    ##Disable Resources##
    "4" {
    
    $name = read-host "Resource to be disabled"

    #check to see if user is on server
        if ((get-mailbox $name) -eq ($null)) {do {write-host "This room was not found"
        $name = read-host -prompt "Enter Room Alias"}
        until ((get-mailbox $name) -ne ($null))}


    disable-mailbox -identity $name
    "
    "
    do {
    $var = read-host -prompt "press x to go back"} until ($var -eq "x")
    if ($var -eq "x") {invoke-expression $scripthome\resources.ps1}
    }
    
    ##Home##
    "x" {Invoke-Expression $scripthome\exchangecontrolpanel.ps1}
   }