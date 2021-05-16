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

$BaseDirectory = "C:\SMM\working\"+"$(Get-Date -Format 'yyyy_MM_dd_HH_mm')_"+$env:UserName
$LogFile = "SMM.log"
$LAP = $BaseDirectory + "\" + $LogFile

$RemoteHost = "sc-exch2016.aslteramo.it"
$OWAHost = "owa.aslteramo.it"
$domainController = "sc-network.aslteramo.it"
$RoutingDomain = "aslteramo.mail.onmicrosoft.com"


New-Item -Path $BaseDirectory -Name $LogFile -ItemType File -Force

# Apertura connessioni OnPrem e OnCloud

#$secpasswd = ConvertTo-SecureString "`$xxxxxxxxxxxxx" -AsPlainText -Force
#$O365Credential = New-Object System.Management.Automation.PSCredential ("xxxxxxxxxxxxx@aslteramo.onmicrosoft.com", $secpasswd)

#$secpasswd = ConvertTo-SecureString "`$xxxxxxxxxxxxxxxx" -AsPlainText -Force
#$OnPremiseCredential = New-Object System.Management.Automation.PSCredential ("aslteramo\xxxxxxxxxxx", $secpasswd)

$OnPremiseCredential = Get-Credential -Message "Inserire le credenziali Exchange On-Premise (aslteramo\username)"
$O365Credential = Get-Credential -Message "Inserire le credenziali Office 365 (username@aslteramo.onmicrosoft.com)"

try {
    $ADSession = New-PSSession -ComputerName $domainController -Credential $OnPremiseCredential
    Invoke-Command $ADSession -Scriptblock { Import-Module ActiveDirectory }
    Import-PSSession -Session $ADSession -module ActiveDirectory -AllowClobber -Prefix Remote

}
catch {
    Log2File -LogFile $LAP -Message "Credenziali AD non valide" -Type "Error"
    return
}

try {
    $ExchSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$($RemoteHost)/PowerShell/" -Authentication Kerberos -Credential $OnPremiseCredential
    Import-PSSession $ExchSession -AllowClobber -Prefix Onprem
}
catch {
    Log2File -LogFile $LAP -Message "Credenziali ExchangeOnPrem non valide" -Type "Error"
    return
}

try {
    $O365Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $O365Credential -Authentication Basic -AllowRedirection
    Import-PSSession $O365Session -AllowClobber -Prefix Online
}
catch {
    Log2File -LogFile $LAP -Message "Credenziali ExchangeOnCloud non valide" -Type "Error"
    return
}

try {
    Import-Module msonline
    Connect-MsolService -Credential $O365Credential
}
catch {
    Log2File -LogFile $LAP -Message "Errore nel modulo per gestire le licenze" -Type "Error"
    return
}


# Avvio Prima Serie di Controlli


if (-Not(Test-Path -Path $InputFile))
{
    Log2File -LogFile $LAP -Message "File di input non esistente" -Type "Error"
    return
}

$mailboxes = Import-Excel -Path $InputFile

Log2File -LogFile $LAP -Message "Avvio controlli verifica input" -Type "Info"


$ValidatedMailboxCollection = @()

foreach ($mailbox in $mailboxes)
{
    $ValidatedMailbox = New-Object -TypeName psobject

    $EmailOnPrem = ([String]$mailbox.emailonprem).Trim()
    $Licenza = ([String]$mailbox.licenza).Trim()

    if ($EmailOnPrem.Length -gt 0)
    {
        Log2File -LogFile $LAP -Message "Avvio controlli verifica mailbox $emailonprem" -Type "Info"
        
        $EmailRegex = '^([\w-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([\w-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$'
        if (-Not($EmailOnPrem -match $EmailRegex))
        {
            Log2File -LogFile $LAP -Message "$EmailOnPrem -> $EmailOnPrem email non valida" -Type "Error"
            return
        }

        if (($Licenza -ne "E3") -and ($Licenza -ne "E1") -and ($Licenza -ne "KIOSK") -and ($Licenza -ne "EOP1"))
        {
            Log2File -LogFile $LAP -Message "$EmailOnPrem -> Licenza $Licenza non valida" -Type "Error"
            return
        }
        

        $mb = Get-OnpremMailbox -Identity $EmailOnPrem
        if (-Not($mb -is [object]))
        {
            Log2File -LogFile $LAP -Message "$EmailOnPrem -> Nessuna mailbox corrispondente a $EmailOnPrem" -Type "Error"
            return
        }

        Clear-Variable -Name mb

        $ValidatedMailbox | Add-Member -MemberType NoteProperty -Name licenza -Value $Licenza
        $ValidatedMailbox | Add-Member -MemberType NoteProperty -Name email -Value $EmailOnPrem

        $ValidatedMailboxCollection += $ValidatedMailbox

    }
}

$ValidatedMailboxCollection


Invoke-Command $ADSession -Scriptblock { repadmin /SyncAll /AedP }
Log2File -LogFile $LAP -Message "Replica AD e Sleep 40 Secondi" -Type "Info"
Start-Sleep -Seconds 40


# MIGRAZIONE MAILBOX IN CLOUD
Log2File -LogFile $LAP -Message "Avvio migrazione mail in cloud" -Type "Info"
foreach ($mailbox in $ValidatedMailboxCollection)
{
    $t = $mailbox.email
    Log2File -LogFile $LAP -Message "Avvio migrazione creazione mailbox $t" -Type "Info"

    try
    {
        New-OnlineMoveRequest -Identity $mailbox.email -RemoteHostName $OWAHost -RemoteCredential $OnPremiseCredential -TargetDeliveryDomain $RoutingDomain -Remote -BadItemLimit 1000 -BatchName $mailbox.email
    }
    catch
    {
        Log2File -LogFile $LAP -Message "Impossibile avviare migrazione" -Type "Info"
        return
    }
	Start-Sleep -Seconds 15
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
    foreach ($mailbox in $ValidatedMailboxCollection) 
    {
        $RequestStatus = (Get-OnlineMoveRequest -BatchName $mailbox.email).Status
        if ($RequestStatus -ne "Completed")
        {
            $emailMigrate = $false
        }
        $message = $mailbox.email+" é nello stato "+$RequestStatus
	    Log2File -LogFile $LAP -Message $message -Type "Info"
	    if ($RequestStatus -eq "Completed") 
	    {
            $IsLicensed = (Get-MsolUser -UserPrincipalName $mailbox.email).IsLicensed
            if (-not $IsLicensed)
            {
                if ($mailbox.licenza -eq "KIOSK")
                {
                    $accountskuid = @()
                    $accountskuid += "aslteramo:EXCHANGEDESKLESS"
                    $accountskuid += "aslteramo:EXCHANGEARCHIVE_ADDON"	
                    $disabledPlans = @()
                    $disabledPlans += "INTUNE_O365"
                    $accountskuid_string = "aslteramo:EXCHANGEDESKLESS"
                }
                if ($mailbox.licenza -eq "E3")
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
                if ($mailbox.licenza -eq "E1")
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
                if ($mailbox.licenza -eq "EOP1")
                {
                    $accountskuid = @()
                    $accountskuid += "aslteramo:EXCHANGESTANDARD"
                    $accountskuid += "aslteramo:EXCHANGEARCHIVE_ADDON"
                    $accountskuid_string = "aslteramo:EXCHANGESTANDARD"
                    $disabledPlans = @()
                    $disabledPlans += "INTUNE_O365"
                }
                $LicenseOptions = New-MsolLicenseOptions -AccountSkuId $accountskuid_string -DisabledPlans $disabledPlans
                Get-MsolUser -UserPrincipalName $mailbox.email | Set-MsolUserLicense -AddLicenses $accountskuid -LicenseOptions $LicenseOptions
                $message = $mailbox.email+" aggiunta licenza "+$mailbox.licenza
	            Log2File -LogFile $LAP -Message $message -Type "Info"
            }
        }           
    }
    Log2File -LogFile $LAP -Message "Sleep 30 secondi" -Type "Info"
    Start-Sleep -Seconds 30
} until ($false)



# AGGIUSTAMENTI POST MIGRAZIONE
Log2File -LogFile $LAP -Message "Settaggi post Migrazione" -Type "Info"
foreach ($mailbox in $ValidatedMailboxCollection) 
{
        Set-OnlineMailbox $mailbox.email -LitigationHoldEnabled $true
        $message = $mailbox.email+" attivazione LegalHold"
        Log2File -LogFile $LAP -Message $message -Type "Info"
	    Set-OnlineMailboxRegionalConfiguration $mailbox.email -TimeZone "W. Europe Standard Time" -Language "it-IT" -DateFormat "dd/MM/yyyy" -TimeFormat "HH:mm" -LocalizeDefaultFolderName
        Set-OnlineMailbox -Identity $mailbox.email -RetentionPolicy "ASL Retention Policy"
        if ($mailbox.licenza -eq "KIOSK")
        {
            $message = $mailbox.email+" attivazione Archive"
            Log2File -LogFile $LAP -Message $message -Type "Info"
            Get-OnpremRemoteMailbox -Identity $mailbox.email | Enable-OnpremRemoteMailbox -Archive
        }

        $id = (Get-OnlineMoveRequest -BatchName $mailbox.email).Guid
        Remove-OnlineMoveRequest -Identity $id -Confirm:$false
}


