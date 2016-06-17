##Group Menu##

$ErrorActionPreference = 'SilentlyContinue'
$scripthome = Get-Location

Do {
Write-Host "
----------Groups----------

1 = View Groups
2 = New Group
3 = Edit Group
4 = Delete Group

x = Home

-------------------------"
"
"
$choice3 = read-host -prompt "Select number & press enter"
} until ($choice3 -eq "1" -or $choice3 -eq "2" -or $choice3 -eq "3" -or $choice3 -eq "4" -or $choice3 -eq "x")

Switch ($choice3) {
    
    ##View Groups## 
    "1" {
    
    $group = read-host -prompt "Enter Name or use * for wildcard"
    
    if ((get-distributiongroup "$group") -eq ($null)) 
        {do {$group = read-host -prompt "That group cannot be found, please try another"}
    until ((get-distributiongroup "$group") -ne ($null))}
    
    Get-distributiongroup $group | FL Name,Primarysmtpaddress
    
    
    do {
    $var = read-host -prompt "press x to go back"} until ($var -eq "x")
    if ($var -eq "x") {invoke-expression $scripthome\groups.ps1}
    }
   
    ##New Group##
    "2" {
    
    $groupname = read-host -prompt "Name of group"
     If ((dsquery group -name $groupname) -ne ($null))
      {do {$groupname = read-host -prompt "This group already exists, please try another"}
    until 
    ((dsquery group -name $groupname) -eq ($null))
    
     }
    
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
  
  
    
      $external = read-host -prompt "Allow External Senders? Y/N?"
     

    If ($external -like "Y*") 
    {new-distributiongroup -name "$groupname" -OrganizationalUnit "$val" | out-null
    set-distributiongroup $groupname -requiresenderauthenticationenabled $false | out-null}
    
   else {new-distributiongroup -name "$groupname" -OrganizationalUnit "$val" | out-null}
   "
   "

   #script uses 'Q' to break, which adds it to the array, which must be removed# 
   $members2 = @()
   Do {$name = read-host -prompt "Enter Alias of member to be added or Q to continue"
   if ($name -ne "Q" -and (get-mailbox $name) -eq ($null)) {"this name cannot be found"
   $name = read-host -prompt "Enter Alias of member to be added or Q to continue"
   $members2 = $members2 + "$name"
   }

   else
   {$members2 = $members2 + "$name"}
   
   
   
   }
   until ($name -eq "Q")
   IF ($name -eq "Q")
   {
    $members = @()
    foreach ($member in $members2)
    {
        if ($member -ne "Q")
        {
            $members = $members += $member
        }
    }
    }
   foreach ($i in $members) {add-distributiongroupmember -identity "$groupname" -member "$i"
   "
   "
   If ((get-mailbox $i) -eq ($null)){write-host "Group Member '$i' was not found"}
    }
   
    get-distributiongroup "$groupname"
    
    write-host 
    "-----Members-----"
    get-distributiongroupmember -identity "$groupname" | ft PrimarySmtpAddress
    "
    "
    do {
    $var = read-host -prompt "press x to go back"} until ($var -eq "x")
    if ($var -eq "x") {invoke-expression $scripthome\groups.ps1}
     
    }
    
    
    ##edit group##
    "3" {Invoke-Expression $scripthome\editgroups.ps1}
    
    ##Delete group##
    "4" {
    $group = read-host -prompt "Enter group name you wish to delete"

    if ((get-distributiongroup $group) -eq ($null)) {do {write-host "This Group was not found"
        $group = read-host -prompt "Enter group name you wish to delete"}
        until ((get-distributiongroup $group) -ne ($null))}




    Remove-Distributiongroup -Identity "$group"
    "
    "
    write-host "$group has been deleted"
    "
    "
    do {
    $var = read-host -prompt "press x to go back"} until ($var -eq "x")
    if ($var -eq "x") {invoke-expression $scripthome\groups.ps1}
    }
    
    
    ##Home##
    "x" {Invoke-Expression $scripthome\exchangecontrolpanel.ps1}
   }