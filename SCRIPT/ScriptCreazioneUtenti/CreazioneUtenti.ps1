[CmdLetBinding()]

Param 
(
	[Parameter(Mandatory=$True)][String]$InputFile
)

function Get-RandomCharacters($length, $characters) {
    $random = 1..$length | ForEach-Object { Get-Random -Maximum $characters.length }
    $private:ofs=""
    return [String]$characters[$random]
}

function Log2File([String]$LogFile,[String]$Message,[ValidateSet('Info','Error')][String]$Type) 
{
	"$($(Get-Date -Format G)) - $($Type): $($Message)" | Out-File -FilePath $LogFile -Append
	if ($Type -eq "Info") {
		Write-Host $Message -ForegroundColor White
	}
	elseif ($Type -eq "Error") {
		Write-Host $Message -ForegroundColor Red
	}
	else {
		Write-Host $Message -ForegroundColor Yellow
	}
}

$BaseDirectory = "C:\SCU\working\"+"$(Get-Date -Format 'yyyy_MM_dd_HH_mm')_"+$env:UserName
#$InputFile = "C:\SCU\TEMPLATE.xlsx"
$LogFile = "SCU.log"
$LAP = $BaseDirectory + "\" + $LogFile

$RemoteHost = "sc-exch2016.aslteramo.it"
$OWAHost = "owa.aslteramo.it"
$domainController = "sc-network.aslteramo.it"
$ADConnectHost = "sc-adconnect.aslteramo.it"
$baseDN_esterni = "OU=__DA_ALLOCARE__,OU=Utenti,DC=aslteramo,DC=it"
$baseDN_interni = "OU=Account Utente e Gruppi,OU=Utenti,DC=aslteramo,DC=it"
$RoutingDomain = "aslteramo.mail.onmicrosoft.com"


$template_dipendenti_oncloud = "C:\SCU\_TEMPLATE_ONCLOUD_.docx"
$template_dipendenti_onprem = "C:\SCU\_TEMPLATE_ONPREM_.docx"

$credenzialiLogFile = $BaseDirectory + "\credenziali.txt"

New-Item -Path $BaseDirectory -Name $LogFile -ItemType File -Force

#$LAP

# Avvio Prima Serie di Controlli


if (-Not(Test-Path -Path $InputFile))
{
    Log2File -LogFile $LAP -Message "File di input non esistente" -Type "Error"
    return
}

$users = Import-Excel -Path $InputFile

Log2File -LogFile $LAP -Message "Avvio controlli verifica input" -Type "Info"


$SamAccountNameCollection = @()
$CFCollection = @()
$EmailAddressCollection = @()
$ValidateUserCollection = @()

foreach ($user in $users)
{
    $ValidatedUser = New-Object -TypeName psobject
    $Cognome = ([String]$user.cognome).Trim()
    $Scadenza = ""
    if ($Cognome.Length -gt 0)
    {
        Log2File -LogFile $LAP -Message "Avvio controlli verifica utente $Cognome" -Type "Info"
        $Nome = ([String]$user.nome).Trim()
        $Username = ([String]$user.username).Trim()
        $Cf = (([String]$user.cf).Trim()).ToUpper();
        $Matricola = ([String]$user.matricola).Trim()
        if ($user.scadenza.Length -gt 0)
        {
            $Scadenza = ($user.scadenza.ToString("dd/MM/yyyy")).Trim()
        }

        $Tipologia = ([String]$user.tipologia).Trim()
        $EmailPrivata = (([String]$user.email_privata).Trim()).ToLower()
        $CreaEmailAsl = ([String]$user.crea_email_asl).Trim()
        $EmailInCloud = ([String]$user.email_in_cloud).Trim()
        $Licenza = ([String]$user.licenza).Trim()
        $EmailCustom = (([String]$user.email_custom).Trim()).ToLower()

        if ($Cognome.Length -le 1) { Log2File -LogFile $LAP -Message "$Cognome -> Cognome non valido" -Type "Error" }
        if ($Nome.Length -le 1) { Log2File -LogFile $LAP -Message "$Cognome -> Nome non valido" -Type "Error" }

        if ($Cf.Length -ne 16) { Log2File -LogFile $LAP -Message "$Cognome -> Codice Fiscale non valido" -Type "Error" }

        if (-Not([String]$Cf.ToCharArray()[0] -match "[A-Z]")) { Log2File -LogFile $LAP -Message "$Cognome -> Codice Fiscale $Cf carattere 1 non valido" -Type "Error" }
        if (-Not([String]$Cf.ToCharArray()[1] -match "[A-Z]")) { Log2File -LogFile $LAP -Message "$Cognome -> Codice Fiscale $Cf carattere 2 non valido" -Type "Error" }
        if (-Not([String]$Cf.ToCharArray()[2] -match "[A-Z]")) { Log2File -LogFile $LAP -Message "$Cognome -> Codice Fiscale $Cf carattere 3 non valido" -Type "Error" }
        if (-Not([String]$Cf.ToCharArray()[3] -match "[A-Z]")) { Log2File -LogFile $LAP -Message "$Cognome -> Codice Fiscale $Cf carattere 4 non valido" -Type "Error" }
        if (-Not([String]$Cf.ToCharArray()[4] -match "[A-Z]")) { Log2File -LogFile $LAP -Message "$Cognome -> Codice Fiscale $Cf carattere 5 non valido" -Type "Error" }
        if (-Not([String]$Cf.ToCharArray()[5] -match "[A-Z]")) { Log2File -LogFile $LAP -Message "$Cognome -> Codice Fiscale $Cf carattere 6 non valido" -Type "Error" }
        if (-Not([String]$Cf.ToCharArray()[6] -match "[0-9]")) { Log2File -LogFile $LAP -Message "$Cognome -> Codice Fiscale $Cf carattere 7 non valido" -Type "Error" }
        if (-Not([String]$Cf.ToCharArray()[7] -match "[0-9]")) { Log2File -LogFile $LAP -Message "$Cognome -> Codice Fiscale $Cf carattere 8 non valido" -Type "Error" }
        if (-Not([String]$Cf.ToCharArray()[8] -match "[A-Z]")) { Log2File -LogFile $LAP -Message "$Cognome -> Codice Fiscale $Cf carattere 9 non valido" -Type "Error" }
        if (-Not([String]$Cf.ToCharArray()[9] -match "[0-9]")) { Log2File -LogFile $LAP -Message "$Cognome -> Codice Fiscale $Cf carattere 10 non valido" -Type "Error" }
        if (-Not([String]$Cf.ToCharArray()[10] -match "[0-9]")) { Log2File -LogFile $LAP -Message "$Cognome -> Codice Fiscale $Cf carattere 11 non valido" -Type "Error" }
        if (-Not([String]$Cf.ToCharArray()[11] -match "[A-Z]")) { Log2File -LogFile $LAP -Message "$Cognome -> Codice Fiscale $Cf carattere 12 non valido" -Type "Error" }
        if (-Not([String]$Cf.ToCharArray()[12] -match "[0-9]")) { Log2File -LogFile $LAP -Message "$Cognome -> Codice Fiscale $Cf carattere 13 non valido" -Type "Error" }
        if (-Not([String]$Cf.ToCharArray()[13] -match "[0-9]")) { Log2File -LogFile $LAP -Message "$Cognome -> Codice Fiscale $Cf carattere 14 non valido" -Type "Error" }
        if (-Not([String]$Cf.ToCharArray()[14] -match "[0-9]")) { Log2File -LogFile $LAP -Message "$Cognome -> Codice Fiscale $Cf carattere 15 non valido" -Type "Error" }
        if (-Not([String]$Cf.ToCharArray()[15] -match "[A-Z]")) { Log2File -LogFile $LAP -Message "$Cognome -> Codice Fiscale $Cf carattere 16 non valido" -Type "Error" }

        $aduser = Get-ADUser -Filter * -Properties SamAccountName,ASLTE-codiceFiscale | Where-Object {$_."ASLTE-codiceFiscale" -eq $Cf} | Select-Object SamAccountName

        if ($aduser -is [object])
        {
            Log2File -LogFile $LAP -Message "$Cognome -> Codice Fiscale $Cf gi� esistente" -Type "Error"
            return
        }
        $CFCollection += $Cf

        if ($Matricola.Length -ne 0)
        {
            if ($Matricola.Length -ne 5) { Log2File -LogFile $LAP -Message "$Cognome -> Matricola $Matricola composta da un numero di caratteri diverso da 5" -Type "Error" }
            if (-Not($Matricola -match "[0-9][0-9][0-9][0-9][0-9]")) { Log2File -LogFile $LAP -Message "$Cognome -> Matricola $Matricola composta da caratteri non validi" -Type "Error" }
        }

        if ($Scadenza.Length -ne 0)
        {
            try
            {
                $s = [datetime]::ParseExact($Scadenza, 'dd/MM/yyyy', $null)
            }
            catch
            {
                Log2File -LogFile $LAP -Message "$Cognome -> Data $Scadenza non valida" -Type "Error"
            }
        }

        if (($Tipologia -ne "Dipendente ASL") -and ($Tipologia -ne "Azienda Esterna")) { Log2File -LogFile $LAP -Message "$Cognome -> Tipologia '$Tipologia' non valida" -Type "Error" } 

        if ($EmailPrivata.Length -eq 0)
        {
            Log2File -LogFile $LAP -Message "$Cognome -> Campo email privata obbligatorio" -Type "Error"
        }
        else
        {
            $EmailRegex = '^([\w-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([\w-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$'
            if (-Not($EmailPrivata -match $EmailRegex))
            {
                Log2File -LogFile $LAP -Message "$Cognome -> $EmailPrivata email non valida" -Type "Error"
                return
            }
        }

        if (($CreaEmailAsl -ne "SI") -and ($CreaEmailAsl -ne "NO"))
        {
            Log2File -LogFile $LAP -Message "$Cognome -> Flag crea email asl non valido" -Type "Error"
            return
        }

        if (($EmailInCloud -ne "SI") -and ($EmailInCloud -ne "NO"))
        {
            Log2File -LogFile $LAP -Message "$Cognome -> Flag email in cloud non valido" -Type "Error"
            return
        } 
        
        if ($Licenza.Length -gt 0)
        {
            if (($Licenza -ne "E3") -and ($Licenza -ne "E1") -and ($Licenza -ne "KIOSK") -and ($Licenza -ne "EOP1"))
            {
                Log2File -LogFile $LAP -Message "$Cognome -> Licenza $Licenza non valida" -Type "Error"
                return
            }
        }

        if (($EmailInCloud -eq "SI") -and ($Licenza.Length -eq 0))
        {
            Log2File -LogFile $LAP -Message "$Cognome -> Email in cloud senza licenza valida" -Type "Error"
            return
        } 

        if ($EmailCustom.Length -ne 0)
        {
            $EmailRegex = '^([\w-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([\w-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$'
            if (-Not($EmailCustom -match $EmailRegex))
            {
                Log2File -LogFile $LAP -Message "$Cognome -> $EmailCustom email custom non valida" -Type "Error"
                return
            }
        }


        # Verifico correttezza SamAccountName
        $SamAccountNameCollection += $Username
        $userExist = $false
        try
        {
            Get-ADUser -Identity $Username
            $userExist = $true
        }
        catch
        {
            $userExist = $false
        }
        if ($userExist)
        {
            Log2File -LogFile $LAP -Message "$Cognome -> Nome utente $Username già in uso" -Type "Error"
            return
        }

        # Generazione indirizzo email e verifica esistenza duplicati

        $EmailAddress = ""
        if ($EmailCustom.Length -ne 0)
        {
            $EmailAddress = $EmailCustom
        }
        else
        {
            $EMLNome = $Nome
            $EMLCognome = $Cognome
            $EMLNome = $EMLNome.Replace("à", "a")
            $EMLNome = $EMLNome.Replace("è", "e")
            $EMLNome = $EMLNome.Replace("é", "e")
            $EMLNome = $EMLNome.Replace("ì", "i")
            $EMLNome = $EMLNome.Replace("ò", "o")
            $EMLNome = $EMLNome.Replace("ù", "u")
            $EMLNome = $EMLNome.Replace(" ", "")
            $EMLNome = $EMLNome.Replace("'", "")
            $EMLCognome = $EMLCognome.Replace("à", "a")
            $EMLCognome = $EMLCognome.Replace("è", "e")
            $EMLCognome = $EMLCognome.Replace("é", "e")
            $EMLCognome = $EMLCognome.Replace("ì", "i")
            $EMLCognome = $EMLCognome.Replace("ò", "o")
            $EMLCognome = $EMLCognome.Replace("ù", "u")
            $EMLCognome = $EMLCognome.Replace(" ", "")
            $EMLCognome = $EMLCognome.Replace("'", "")
            $EMLNome = $EMLNome.ToLower()
            $EMLCognome = $EMLCognome.ToLower()
            $EmailAddress = $EMLNome+"."+$EMLCognome+"@aslteramo.it"
        }

        $UPN = $EmailAddress.Replace("@aslteramo.it","")


        $aduser = Get-ADUser -Filter * | Where-Object {$_."UserPrincipalName" -eq $EmailAddress} | Select-Object SamAccountName
        if ($aduser -is [object])
        {
            Log2File -LogFile $LAP -Message "$Cognome -> UPN $EmailAddress già in uso, è necessario popolare la colonna EMAIL_CUSTOM anche nel caso in cui non è necessario creare la cassetta di posta elettronica" -Type "Error"
            return
        }
        $EmailAddressCollection += $EmailAddress

        $ValidatedUser | Add-Member -MemberType NoteProperty -Name cognome -Value $Cognome
        $ValidatedUser | Add-Member -MemberType NoteProperty -Name nome -Value $Nome
        $ValidatedUser | Add-Member -MemberType NoteProperty -Name username -Value $Username
        $ValidatedUser | Add-Member -MemberType NoteProperty -Name cf -Value $Cf
        $ValidatedUser | Add-Member -MemberType NoteProperty -Name matricola -Value $Matricola
        $ValidatedUser | Add-Member -MemberType NoteProperty -Name scadenza -Value $Scadenza
        $ValidatedUser | Add-Member -MemberType NoteProperty -Name tipologia -Value $Tipologia
        $ValidatedUser | Add-Member -MemberType NoteProperty -Name email_privata -Value $EmailPrivata
        $ValidatedUser | Add-Member -MemberType NoteProperty -Name crea_email_asl -Value $CreaEmailAsl
        $ValidatedUser | Add-Member -MemberType NoteProperty -Name email_in_cloud -Value $EmailInCloud
        $ValidatedUser | Add-Member -MemberType NoteProperty -Name licenza -Value $Licenza
        $ValidatedUser | Add-Member -MemberType NoteProperty -Name email_asl -Value $EmailAddress

        $ValidateUserCollection += $ValidatedUser

    }
}

# Verifico le collection alla ricerca di duplicati

Log2File -LogFile $LAP -Message "Avvio controlli presenza duplicati" -Type "Info"

$tmp = @()
$tmp = $SamAccountNameCollection | Select-Object -Unique
if ($SamAccountNameCollection.Count -ne $tmp.Count)
{
    Log2File -LogFile $LAP -Message "Ci sono dei duplicati nell'elenco dei nomi utente" -Type "Error"
    $SamAccountNameCollection
    return    
}

$tmp = @()
$tmp = $CFCollection | Select-Object -Unique
if ($CFCollection.Count -ne $tmp.Count)
{
    Log2File -LogFile $LAP -Message "Ci sono dei duplicati nell'elenco dei codici fiscali" -Type "Error"
    $CFCollection
    return    
}

$tmp = @()
$tmp = $EmailAddressCollection | Select-Object -Unique
if ($EmailAddressCollection.Count -ne $tmp.Count)
{
    Log2File -LogFile $LAP -Message "Ci sono dei duplicati nell'elenco dell'email" -Type "Error"
    $EmailAddressCollection
    return    
}


#$secpasswd = ConvertTo-SecureString "`$`$xxxxxxxxxxxxxx" -AsPlainText -Force
#$O365Credential = New-Object System.Management.Automation.PSCredential ("xxxxxxxxxxxxx@xxxxxxxxxx.onmicrosoft.com", $secpasswd)

#$secpasswd = ConvertTo-SecureString "`$`$xxxxxxxxxxxxxxxxx" -AsPlainText -Force
#$OnPremiseCredential = New-Object System.Management.Automation.PSCredential ("xxxxxxxxx\yyyyyyyyy", $secpasswd)

$OnPremiseCredential = Get-Credential -Message "Inserire le credenziali Exchange On-Premise (aslteramo\username)"
$O365Credential = Get-Credential -Message "Inserire le credenziali Office 365 (username@aslteramo.onmicrosoft.com)"

try
{
    $ADSession = New-PSSession -ComputerName $domainController -Credential $OnPremiseCredential
    Invoke-Command $ADSession -Scriptblock { Import-Module ActiveDirectory }
    Import-PSSession -Session $ADSession -module ActiveDirectory -AllowClobber -Prefix Rem
}
catch
{
    Log2File -LogFile $LAP -Message "Credenziali Errate per sessione AD" -Type "Error"
    return
}

try
{
    $ExchSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$($RemoteHost)/PowerShell/" -Authentication Kerberos -Credential $OnPremiseCredential
    Import-PSSession $ExchSession -AllowClobber -Prefix Onprem
}
catch
{
    Log2File -LogFile $LAP -Message "Credenziali Errate per sessione Exchange OnPrem" -Type "Error"
    return
}

try
{
    $O365Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $O365Credential -Authentication Basic -AllowRedirection
    Import-PSSession $O365Session -AllowClobber -Prefix Online
}
catch
{
    Log2File -LogFile $LAP -Message "Credenziali Errate per sessione Exchange OnCloud" -Type "Error"
    return    
}

try
{
    Import-Module msonline
    Connect-MsolService -Credential $O365Credential
}
catch
{
    Log2File -LogFile $LAP -Message "Errore caricamento modulo gestione licenze" -Type "Error"
    return     
}


# AVVIO CREAZIONE UTENTI

Log2File -LogFile $LAP -Message "Avvio creazione utenti" -Type "Info"


foreach ($user in $ValidateUserCollection)
{
    $passwd1 = Get-RandomCharacters -length 6 -characters 'abcdefghijklmnopqrstuvwxyz'
    $passwd2 = Get-RandomCharacters -length 1 -characters '1234567890'
    $passwd3 = Get-RandomCharacters -length 1 -characters 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    $passwd = $passwd1 + $passwd2 + $passwd3
    $secpasswd = ConvertTo-SecureString $passwd -AsPlainText -Force
    $ADNAME = $user.cognome+" "+$user.nome
    if ($user.tipologia -eq "Dipendente ASL")
    {
        $baseDN = $baseDN_interni   
    }
    else
    {
        $baseDN = $baseDN_esterni
    }
    New-RemADUser -Path $baseDN -SamAccountName $user.username -GivenName $user.nome -Surname $user.cognome -Name $ADNAME -DisplayName $ADNAME -UserPrincipalName $user.email_asl -AccountPassword $secpasswd -ChangePasswordAtLogon $true -Enabled $true
    
    Add-RemADGroupMember -Identity GRP_WWW_utente_standard -Members $user.username -Server $domainController
    Add-RemADGroupMember -Identity GRP_POLICY-PWD_Dipendenti -Members $user.username -Server $domainController
    #Add-RemADGroupMember -Identity GRP_POLICY-PWD_VPN_User -Members $user.username -Server $domainController
    Set-RemADUser -Identity $user.username -Add @{'ASLTE-codiceFiscale'=$user.cf} -Server $domainController
    if ($user.matricola.Legth -gt 0)
    {
        Set-RemADUser -Identity $user.username -Replace @{EmployeeID=$user.matricola} -Server $domainController
    }
    if ($user.scadenza.Length -gt 0)
    {
        Set-RemADAccountExpiration -Identity $user.username -DateTime $user.scadenza -Server $domainController
    }
    Log2File -LogFile $LAP -Message "Creato utente $ADNAME" -Type "Info"
    if ($user.crea_email_asl -ne "SI")
    {
        Set-RemADUser -Identity $user.username -EmailAddress $user.email_privata
    }

    
    $line = $user.username+","+$passwd+","+$user.email_privata
    $line | Out-File -FilePath $credenzialiLogFile -Append


    if ($user.crea_email_asl -eq "SI")
    {
        if ($user.email_in_cloud -eq "SI")
        {
            $template_dipendenti = $template_dipendenti_oncloud
        }
        else
        {
            $template_dipendenti = $template_dipendenti_onprem
        }
        # CREAZIONE FILE WORD PER RILASCIO CREDENZIALI
        $objWord = New-Object -comobject Word.Application  
        $objWord.Visible = $false
        $objDoc = $objWord.Documents.Open($template_dipendenti, $false, $true)

        $objSelection = $objWord.Selection 

        $wdFindContinue = 1
        $MatchCase = $true 
        $MatchWholeWord = $true
        $MatchWildcards = $False 
        $MatchSoundsLike = $False 
        $MatchAllWordForms = $False 
        $Forward = $True 
        $Wrap = $wdFindContinue 
        $Format = $False 
        $wdReplaceNone = 0
        $wdReplaceAll = $true

        $FindText = "{username}"
        $ReplaceWith = $user.username
        $a = $objSelection.Find.Execute($FindText,$MatchCase,$MatchWholeWord,$MatchWildcards,$MatchSoundsLike,$MatchAllWordForms,$Forward,$Wrap,$Format,$ReplaceWith,$wdReplaceAll)

        $FindText = "{password}"
        $ReplaceWith = $passwd
        $a = $objSelection.Find.Execute($FindText,$MatchCase,$MatchWholeWord,$MatchWildcards,$MatchSoundsLike,$MatchAllWordForms,$Forward,$Wrap,$Format,$ReplaceWith,$wdReplaceAll)


        $FindText = "{email}"
        $ReplaceWith = $user.email_asl
        $a = $objSelection.Find.Execute($FindText,$MatchCase,$MatchWholeWord,$MatchWildcards,$MatchSoundsLike,$MatchAllWordForms,$Forward,$Wrap,$Format,$ReplaceWith,$wdReplaceAll)

        $tmp = $user.email_privata.Replace("@","_AT_")
        $tmp = $tmp.Replace(".","_DOT_")


        $wordFileName = $BaseDirectory + "\" + $user.username + " - credenziali.pdf"

        $objDoc.SaveAs($wordFileName,17)
        $objDoc.Close($false)
        $objWord.Quit()
    }

}

Invoke-Command $ADSession -Scriptblock { repadmin /SyncAll /AedP }
Log2File -LogFile $LAP -Message "Replica AD e Sleep 10 Secondi" -Type "Info"
Start-Sleep -Seconds 10


# AVVIO CREAZIONE MAILBOX ONPREM

Log2File -LogFile $LAP -Message "Avvio creazione mailbox on prem" -Type "Info"
foreach ($user in $ValidateUserCollection)
{
    if ($user.crea_email_asl -eq "SI")
    {
        $ADNAME = $user.cognome+" "+$user.nome
        $alias = ""
        $alias = $user.email_asl.Replace("@aslteramo.it","")
        Log2File -LogFile $LAP -Message "Avvio creazione mailbox $alias" -Type "Info"
        Enable-OnpremMailbox -Identity $user.username -Alias $alias -DisplayName $ADNAME -Database MBX01
        Start-Sleep -Seconds 20
        Set-OnpremMailboxRegionalConfiguration $alias -TimeZone "W. Europe Standard Time" -Language "it-IT" -DateFormat "dd/MM/yyyy" -TimeFormat "HH:mm" -LocalizeDefaultFolderName
        Get-OnpremMailbox -Identity $alias | Set-OnpremMailbox -DomainController $domainController -EmailAddressPolicyEnabled $false
    }
}

Invoke-Command $ADSession -Scriptblock { repadmin /SyncAll /AedP }
Log2File -LogFile $LAP -Message "Replica AD e Sleep 20 Secondi" -Type "Info"
Start-Sleep -Seconds 20

Invoke-Command $ADSession -Scriptblock { repadmin /SyncAll /AedP }
Log2File -LogFile $LAP -Message "Replica AD e Sleep 20 Secondi" -Type "Info"
Start-Sleep -Seconds 20


$ADConnectSession = New-PSSession -ComputerName $ADConnectHost -Credential $OnPremiseCredential
Invoke-Command $ADConnectSession -Scriptblock { Start-ADSyncSyncCycle -Policy Delta }
Log2File -LogFile $LAP -Message "Replica ADConnect e Sleep 400 Secondi" -Type "Info"
Start-Sleep -Seconds 400

# MIGRAZIONE MAILBOX IN CLOUD
Log2File -LogFile $LAP -Message "Avvio migrazione mail in cloud" -Type "Info"
foreach ($user in $ValidateUserCollection)
{
    if ($user.crea_email_asl -eq "SI" -and $user.email_in_cloud -eq "SI")
    {
        $t = $user.email_asl
        Log2File -LogFile $LAP -Message "Avvio migrazione creazione mailbox $t" -Type "Info"
        try
        {
            New-OnlineMoveRequest -Identity $user.email_asl -RemoteHostName $OWAHost -RemoteCredential $OnPremiseCredential -TargetDeliveryDomain $RoutingDomain -Remote -BadItemLimit 1000 -BatchName $user.email_asl
        }
        catch
        {
           Log2File -LogFile $LAP -Message "Impossibile avviare migrazione" -Type "Info"
           return
        }
	    Start-Sleep -Seconds 15
    }
}

# GESTIONE E FINALIZZAZIONE JOB DI MIGRAZIONE
Log2File -LogFile $LAP -Message "Gestione JOB di Migrazione" -Type "Info"
$emailMigrate = $false
do
{
    if ($emailMigrate)
    {
        break
    }
    $emailMigrate = $true
	foreach ($user in $ValidateUserCollection) 
	{
		if ($user.crea_email_asl -eq "SI" -and $user.email_in_cloud -eq "SI")
	    {
            $RequestStatus = (Get-OnlineMoveRequest -BatchName $user.email_asl).Status
            if ($RequestStatus -ne "Completed")
            {
                $emailMigrate = $false
            }
            $message = $user.email_asl+" é nello stato "+$RequestStatus
			Log2File -LogFile $LAP -Message $message -Type "Info"
			if ($RequestStatus -eq "Completed"-or $RequestStatus -eq "CompletedWithWarning") 
			{
                $IsLicensed = (Get-MsolUser -UserPrincipalName $user.email_asl).IsLicensed
                if (-not $IsLicensed)
                {
                    if ($user.licenza -eq "KIOSK")
                    {
	                    $accountskuid = @()
	                    $accountskuid += "aslteramo:EXCHANGEDESKLESS"
	                    $accountskuid += "aslteramo:EXCHANGEARCHIVE_ADDON"	
	                    $disabledPlans = @()
	                    $disabledPlans += "INTUNE_O365"
	                    $accountskuid_string = "aslteramo:EXCHANGEDESKLESS"
                    }

                    if ($user.licenza -eq "E3")
                    {
	                    $accountskuid = "aslteramo:ENTERPRISEPACK"
	                    $accountskuid_string = "aslteramo:ENTERPRISEPACK"
	                    $disabledPlans = @()
	                    $disabledPlans += "KAIZALA_O365_P3" 
	                    $disabledPlans += "MICROSOFT_SEARCH"
	                    $disabledPlans += "WHITEBOARD_PLAN2"
	                    $disabledPlans += "MIP_S_CLP1"
	                    $disabledPlans += "MYANALYTICS_P2"
	                    $disabledPlans += "BPOS_S_TODO_2"
	                    $disabledPlans += "FORMS_PLAN_E3"
	                    $disabledPlans += "STREAM_O365_E3"
	                    $disabledPlans += "Deskless"
	                    $disabledPlans += "FLOW_O365_P2"
	                    $disabledPlans += "POWERAPPS_O365_P2"
	                    $disabledPlans += "TEAMS1"
	                    $disabledPlans += "PROJECTWORKMANAGEMENT"
	                    $disabledPlans += "SWAY"
	                    $disabledPlans += "INTUNE_O365"
	                    $disabledPlans += "YAMMER_ENTERPRISE"
	                    $disabledPlans += "RMS_S_ENTERPRISE"
	                    $disabledPlans += "OFFICESUBSCRIPTION"
	                    $disabledPlans += "MCOSTANDARD"
	                    $disabledPlans += "SHAREPOINTWAC"
	                    $disabledPlans += "SHAREPOINTENTERPRISE"
                    }

                    if ($user.licenza -eq "E1")
                    {
	                    $accountskuid = @()
	                    $accountskuid += "aslteramo:STANDARDPACK"
	                    $accountskuid += "aslteramo:EXCHANGEARCHIVE_ADDON"
	                    $accountskuid_string = "aslteramo:STANDARDPACK"
	                    $disabledPlans = @()
	                    $disabledPlans += "KAIZALA_O365_P2" 
	                    $disabledPlans += "MICROSOFT_SEARCH"
	                    $disabledPlans += "WHITEBOARD_PLAN1"
	                    $disabledPlans += "MYANALYTICS_P2"
	                    $disabledPlans += "BPOS_S_TODO_1"
	                    $disabledPlans += "FORMS_PLAN_E1"
	                    $disabledPlans += "STREAM_O365_E1"
	                    $disabledPlans += "Deskless"
	                    $disabledPlans += "FLOW_O365_P1"
	                    $disabledPlans += "POWERAPPS_O365_P1"
	                    $disabledPlans += "TEAMS1"
	                    $disabledPlans += "PROJECTWORKMANAGEMENT"
	                    $disabledPlans += "SWAY"
	                    $disabledPlans += "INTUNE_O365"
	                    $disabledPlans += "YAMMER_ENTERPRISE"
	                    $disabledPlans += "MCOSTANDARD"
	                    $disabledPlans += "SHAREPOINTWAC"
	                    $disabledPlans += "OFFICEMOBILE_SUBSCRIPTION"
	                    $disabledPlans += "SHAREPOINTSTANDARD"
                    }
                    if ($user.licenza -eq "EOP1")
                    {
                        $accountskuid = @()
                        $accountskuid += "aslteramo:EXCHANGESTANDARD"
                        $accountskuid += "aslteramo:EXCHANGEARCHIVE_ADDON"
                        $accountskuid_string = "aslteramo:EXCHANGESTANDARD"
                        $disabledPlans = @()
                        $disabledPlans += "INTUNE_O365"
                    }

                    $LicenseOptions = New-MsolLicenseOptions -AccountSkuId $accountskuid_string -DisabledPlans $disabledPlans
                    Get-MsolUser -UserPrincipalName $user.email_asl | Set-MsolUserLicense -AddLicenses $accountskuid -LicenseOptions $LicenseOptions
                    $message = $user.email_asl+" aggiunta licenza "+$user.licenza
			        Log2File -LogFile $LAP -Message $message -Type "Info"
                }
            }            
		}
	}
	Log2File -LogFile $LAP -Message "Sleep 30 secondi" -Type "Info"
	sleep -Seconds 30
} until ($false)




# AGGIUSTAMENTI POST MIGRAZIONE
Log2File -LogFile $LAP -Message "Settaggi post Migrazione" -Type "Info"
foreach ($user in $ValidateUserCollection) 
{
    if ($user.crea_email_asl -eq "SI" -and $user.email_in_cloud -eq "SI")
    {
        Set-OnlineMailbox $user.email_asl -LitigationHoldEnabled $true
        $message = $user.email_asl+" attivazione LegalHold"
        Log2File -LogFile $LAP -Message $message -Type "Info"
	    Set-OnlineMailboxRegionalConfiguration $user.email_asl -TimeZone "W. Europe Standard Time" -Language "it-IT" -DateFormat "dd/MM/yyyy" -TimeFormat "HH:mm" -LocalizeDefaultFolderName
        Set-OnlineMailbox -Identity $user.email_asl -RetentionPolicy "ASL Retention Policy"
        if ($user.licenza -eq "KIOSK")
        {
            $message = $user.email_asl+" attivazione Archive"
            Log2File -LogFile $LAP -Message $message -Type "Info"
            Get-OnpremRemoteMailbox -Identity $user.email_asl | Enable-OnpremRemoteMailbox -Archive
        }

        $id = (Get-OnlineMoveRequest -BatchName $user.email_asl).Guid
        Remove-OnlineMoveRequest -Identity $id -Confirm:$false
    }
}


