<#
.SYNOPSIS
  Dieses Script installiert ein LetsEncrypt Zertifikat mittels dns-01 und bindet dieses in einige Exchange Server Dienste ein.
  Die dns-01 Methode benötigt in diesem Fall, dass die Domäne sich mit den Nameservern bei Hetzner-DNS befindet. 
.DESCRIPTION
  Für die Benutzung wird das Powershell Modul "PoSh-ACME" benötigt. Das Script installiert die neueste Version automatisch, falls keine vorhanden ist. 
  Das Script sollte außerdem nur in der Exchange Management Konsole ausgeführt werden. 
.OUTPUTS
  Die aktuelle Zertifikatsdatei wird unter $certFilePath auf der ausführenden Maschine und unter "C:\cert.pfx" auf allen Exchange Servern abgelegt. 
.NOTES
  Version:          1.0
  Autor:            Tobias Hertel
  Erstelldatum:     13.03.2023
  Zweck:            Erstmalige Erstellung des Scripts
#>

#edit the following variables
#path for hoarding all certificates
$certFilePath = "C:\cert\exchange_$(get-date -format yyyy-MM-dd).pfx"
#your domain names
$certNames = 'mail.example.com','autodiscover.example.com'
#your email address
$certMail = 'mail@example.com'
#your exchange servers
$exchangeServers = "exchange01.example.com","exchange02.example.com"
#your hetzner-dns api key
$hetznerAPI = 'xxxxxx'
#stop editing from here

#password for privatekey
$certPw = ConvertTo-SecureString -String "poshacme" -Force -AsPlainText

#Check if Module is available
if(((Get-Module -ListAvailable | Where-Object {$_.Name -eq "PoSh-ACME"}).count) -ge 1) {
    #if availabe, import
    Import-Module PoSh-ACME
} else {
    #if not available, install and import
    Install-Module -Name Posh-ACME -Scope CurrentUser
    Import-Module PoSh-ACME
}

#set ACME-Server 
#LE_STAGE = LetsEncrypt testing
#LE_PROD = LetsEncrypt productive
Set-PAServer LE_STAGE

#if first time using PoSh-ACME, create ACME-Account with mail, accept terms of use and set private key to rsa
if(!((Get-PAAccount -List).id.count) -ge 1){
    New-PAAccount -AcceptTOS -Contact $certMail -KeyLength 4096
}

#if no cert was yet requested
if((Get-PACertificate $certNames[0]).Subject.count -eq 0){
    #set Plugin Args for HetznerDNS
    $pArgs = @{
        HetznerToken = ConvertTo-SecureString -AsPlainText -Force -String $hetznerAPI
    }
    #request new certificate
    New-PACertificate $certNames -Plugin Hetzner -PluginArgs $pArgs -AcceptTOS -Contact $certMail
}else {
    #renew current certificate
    Submit-Renewal -MainDomain $certNames[0]
}

#get requested certificate
$cert = Get-PACertificate -MainDomain $certNames[0]

#copy new certificate in either a temp folder or a cert database
Copy-Item -Path $cert.PfxFile -Destination $certFilePath

#get newest Cert FileData
$certData = Get-Content -Encoding byte -ReadCount 0 $certFilePath

#enable new certificate on all servers
foreach ($exchangeServer in $exchangeServers) {
    #import new certificate on each exchange server
    Import-ExchangeCertificate -Server $exchangeServer -FileData $certData -Password $certPw

    #enable new certificate on each exchange server for services POP3,IMAP,SMTP and IIS Site
    Enable-ExchangeCertificate -Server $exchangeServer -Thumbprint $cert.Thumbprint -Services POP,IMAP,IIS,SMTP -Force
}
