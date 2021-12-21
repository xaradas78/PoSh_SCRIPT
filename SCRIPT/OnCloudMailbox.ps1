# Per ogni mailbox in cloud scrive nel file $outputFile diverse informazioni


$outputFile = ".\oncloudmailbox.txt"

Remove-Item -Path $outputFile -Force -Confirm:$false


$UserCredential = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell-liveid?DelegatedOrg=ASLTERAMO.IT -Credential $UserCredential -Authentication Basic -AllowRedirection
Import-PSSession $Session


$mailboxes = Get-Mailbox -ResultSize Unlimited | Select-Object UserPrincipalName
#$mailboxes = Get-Mailbox -ResultSize 10 | Select-Object UserPrincipalName

foreach ($mb in $mailboxes)
{
    $upn = $mb.UserPrincipalName
    $user = Get-ADUser -Filter "userPrincipalName -eq '$upn'"  -Properties * | Select-Object SamAccountName,AccountExpirationDate, ASLTE-codiceFiscale, givenName, sn

    $accountExpirationDate = $user.AccountExpirationDate
    if ($null -eq $accountExpirationDate.AccountExpirationDate)
    {
        $accountExpirationDate = "NEVER"
    }
    else
    {
        $accountExpirationDate = $accountExpirationDate.AccountExpirationDate.ToString()
    }
    $mail = $mb.UserPrincipalName
    $string = $user.'ASLTE-codiceFiscale'+","+$user.sn+","+$user.givenName+","+ $user.samAccountName+","+$mail+","+$accountExpirationDate
    Add-Content $outputFile $string
}
