$usercredential = get-credential

$server = hostname

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$server/PowerShell/ -Authentication Kerberos -Credential $UserCredential
import-pssession $session

$scripthome = Get-Location

invoke-expression $scripthome\exchangecontrolpanel.ps1