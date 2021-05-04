function Log2File([String]$LogFile,[String]$Message,[ValidateSet('Info','Warning','Error')][String]$Type) 
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

# Variabili
[string]$ApipaAddress = "169.254"
[string]$BaseDirectory = "C:\SWSIA\logs"
[string]$LogFile = "IpResetOnDhcpFail.log"
[string]$LAP = $BaseDirectory + "\" + $LogFile
[int]$maxLogFileSizeKB = 15000


# Logica
if(!(Test-Path -path $BaseDirectory))  
{  
    New-Item -ItemType directory -Path $BaseDirectory
}

[int]$logFileSize = [math]::Round((Get-Item $LAP).length/1024)

if ($logFileSize -gt $maxLogFileSizeKB)
{
    Remove-Item -Path $LAP -Force
}

Log2File -LogFile $LAP -Message "Avvio controllo" -Type "Info"

$currentConf = Get-NetIPConfiguration -Detailed | Out-String -Width 200
Log2File -LogFile $LAP -Message "Configurazione corrente" -Type "Info"
Log2File -LogFile $LAP -Message $currentConf -Type "Info"
$currentConf = ipconfig /all | Out-String -Width 200
Log2File -LogFile $LAP -Message $currentConf -Type "Info"

$ipAddress = Get-NetAdapter -Physical | Where-Object {$_.Status -eq "Up"} | Get-NetIPAddress -AddressFamily IPv4 | Select-Object IPAddress

[string]$ip = $ipAddress.IPAddress.Substring(0,7)

Log2File -LogFile $LAP -Message "Valore ip: $ip" -Type "Info"

#$ip = "169.254"

if ($ip -eq $ApipaAddress)
{
    Log2File -LogFile $LAP -Message "Attenzione IP in APIPA" -Type "Error"

    Log2File -LogFile $LAP -Message "Avvio azione correttiva: netsh winsock reset" -Type "Warning"
    netsh winsock reset

    Log2File -LogFile $LAP -Message "Avvio azione correttiva: netsh int ip reset" -Type "Warning"
    netsh int ip reset

    Log2File -LogFile $LAP -Message "Avvio azione correttiva: netsh advfirewall reset" -Type "Warning"
    netsh advfirewall reset

    Log2File -LogFile $LAP -Message "Avvio azione correttiva: netsh winhttp>reset proxy" -Type "Warning"
    netsh winhttp>reset proxy

    Log2File -LogFile $LAP -Message "Avvio azione correttiva: ipconfig /release" -Type "Warning"
    ipconfig /release

    Log2File -LogFile $LAP -Message "Avvio azione correttiva: ipconfig /renew" -Type "Warning"
    ipconfig /renew

    Log2File -LogFile $LAP -Message "Avvio azione correttiva: ipconfig /flushdns" -Type "Warning"
    ipconfig /flushdns
}
else
{
    Log2File -LogFile $LAP -Message "DHCP sembra essere operativo" -Type "Info"   
}