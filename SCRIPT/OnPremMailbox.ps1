# Per ogni mailbox onprem scrive nel file $outputFile diverse informazioni

$exchangeServer = "sc-exch2016"
$outputFile = ".\onpremmailbox.txt"

Remove-Item -Path $outputFile -Force -Confirm:$false

$UserCredential = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$exchangeServer/PowerShell/ -Authentication Kerberos -Credential $UserCredential
Import-PSSession $Session -DisableNameChecking


$mailboxes = Get-Mailbox -ResultSize unlimited

foreach ($mb in $mailboxes)
{
    $samAccountName = $mb.SamAccountName
    $user = Get-ADUser -Identity $mb.SamAccountName -Properties * | Select-Object AccountExpirationDate, ASLTE-codiceFiscale, givenName, sn
    $accountExpirationDate = $user.AccountExpirationDate
    if ($null -eq $accountExpirationDate.AccountExpirationDate)
    {
        $accountExpirationDate = "NEVER"
    }
    else
    {
        $accountExpirationDate = $accountExpirationDate.AccountExpirationDate.ToString()
    }
    $mail = $mb.PrimarySmtpAddress
    $string = $user.'ASLTE-codiceFiscale'+","+$user.sn+","+$user.givenName+","+$samAccountName+","+$mail+","+$accountExpirationDate
    Add-Content $outputFile $string
}
