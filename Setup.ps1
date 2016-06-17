##Setup##


$scripthome = Get-Location

Do {
Write-Host "
----------Setup----------

 
1 = Required Permissions



x = Home

---------------------------"
"
"
$choice8 = read-host -prompt "Select number & press enter"
"
"
} until ($choice8 -eq "1" -or $choice8 -eq "2" -or $choice8 -eq "3" -or $choice8 -eq "4" -or $choice8 -eq "x")

Switch ($choice8) 
{
    
    <##Universal Variables## 
    "1" {}#>
    ##Required Permissions##
    "1"{
    ""
     
    write-host "This will add the Reset Password role to the Organization management role group, press Enter to continue"
    $x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

    new-managementroleassignment -securitygroup "Organization Management" -Role "Reset Password" 
    "
    "
    $R = read-host -prompt "Would you like to add another user to the Organization Management role group Y/N?"
    ""
    if ($R -like "y*")
        {
        $member = read-host -prompt "Enter username to be placed in Organization Management role group, Administrator is default"
        Add-RolegroupMember "Organization Management" -Member $member
        }
        ""
        
        write-host "Current Members:"

        get-rolegroupmember "Organization management" | ft Alias
        "
        " 
        
   
    do {$var = read-host -prompt "press x to go back"} until ($var -eq "x")
    if ($var -eq "x") {invoke-expression $scripthome\setup.ps1}
 
 }
    ##Home##
    "x" {Invoke-Expression $scripthome\exchangecontrolpanel.ps1}
}