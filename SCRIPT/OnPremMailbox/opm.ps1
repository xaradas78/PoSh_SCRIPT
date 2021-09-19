# Per ogni mailbox del server $exchangeServer scrive nel file $outputFile il SamAccountName
# l'indirizzo di posta e la data di scadenza dell'account

$exchangeServer = "sc-exch2016"
$outputFile = ".\test.txt"

Remove-Item -Path $outputFile -Force -Confirm:$false

$UserCredential = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$exchangeServer/PowerShell/ -Authentication Kerberos -Credential $UserCredential
Import-PSSession $Session -DisableNameChecking


Add-Content $outputFile "SamAccountName,PrimarySmtpAddress,AccountExpirationDate"

$mailboxes = Get-Mailbox -ResultSize unlimited

foreach ($mb in $mailboxes)
{
    $samAccountName = $mb.SamAccountName
    $accountExpirationDate = Get-ADUser -Identity $mb.SamAccountName -Properties * | Select-Object AccountExpirationDate
    if ($null -eq $accountExpirationDate.AccountExpirationDate)
    {
        $accountExpirationDate = "NEVER"
    }
    else
    {
        $accountExpirationDate = $accountExpirationDate.AccountExpirationDate.ToString()
    }
    $mail = $mb.PrimarySmtpAddress
    $string = $samAccountName+","+$mail+","+$accountExpirationDate
    Add-Content $outputFile $string
}