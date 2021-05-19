#$UserCredential = Get-Credential
#$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://sc-exch2016/PowerShell/ -Authentication Kerberos -Credential $UserCredential
#Import-PSSession $Session -DisableNameChecking

$inputFile = ".\mailbox.txt"
$outputFile = ".\output.txt"
$attribute = @("msExchMailboxMoveBatchName","msExchMailboxMoveFlags","msExchMailboxMoveRemoteHostName","msExchMailboxMoveSourceArchiveMDBLink","msExchMailboxMoveSourceArchiveMDBLinkSL","msExchMailboxMoveSourceMDBLink","msExchMailboxMoveSourceMDBLinkSL","msExchMailboxMoveStatus","msExchMailboxMoveTargetArchiveMDBLink","msExchMailboxMoveTargetArchiveMDBLinkSL","msExchMailboxMoveTargetMDBLink","msExchMailboxMoveTargetMDBLinkSL")


foreach($line in Get-Content $inputFile)
{
    $mbx = Get-Mailbox -Identity $line
    if ($mbx.MailboxMoveBatchName.Length -gt 0)
    {
        $string = $line + "#" + $mbx.Id
        #Write-Host $string
        Add-Content -Path $outputFile -Value $string
    }
}


