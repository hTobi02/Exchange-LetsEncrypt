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

  #your email address
  $certMail = 'mail@example.com'

  #your domain names
  $certNames = 'mail.example.com','autodiscover.example.com'
  #if you need different plugins for multiple domains
  # $certNames = 'mail.example.com','mail.example2.com','autodiscover.example.com','autodiscover.example2.com'

  #your plugin names
  #all plugins: AcmeDns,Active24,Akamai,Aliyun,All-Inkl,Aurora,AutoDNS,Azure,Beget,BlueCat,Bunny,Cloudflare,ClouDNS,Combell,Constellix,CoreNetworks,DeSEC,DMEasy,DNSimple,DNSPod,DOcean,DomainOffensive,Domeneshop,Dreamhost,DuckDNS,Dynu,EasyDNS,Easyname,FreeDNS,Gandi,GCloud,Google,GoDaddy,Hetzner,HostingDe,HurricaneElectric,IBMSoftLayer,Infoblox,Infomaniak,IONOS,ISPConfig,LeaseWeb,Linode,Loopia,LuaDns,Manual,Namecheap,NameCom,NS1,OVH,PointDNS,Porkbun,PortsManagement,Rackspace,Regru,RFC2136,Route53,Selectel,SimpleDNSPlus,Simply,SSHProxy,TotalUptime,UKFast,Windows,Yandex,Zilore,Zonomi
  #read more here: https://poshac.me/docs/v4/Plugins/
  $plugins = "Hetzner,Hetzner"
  #if you need different plugins for multiple domains
  # $plugins = "Hetzner,Cloudflare,Hetzner,Cloudflare"

  #your plugin api keys (Uncomment for your needs)
  # $token_AcmeDns = 'xxxxxx'
  # $token_Active24 = 'xxxxxx'
  # $token_Akamai = 'xxxxxx'
  # $token_Aliyun = 'xxxxxx'
  # $token_All-Inkl = 'xxxxxx'
  # $token_Aurora = 'xxxxxx'
  # $token_AutoDNS = 'xxxxxx'
  # $token_Azure = 'xxxxxx'
  # $token_Beget = 'xxxxxx'
  # $token_BlueCat = 'xxxxxx'
  # $token_Bunny = 'xxxxxx'
  # $token_Cloudflare = 'xxxxxx'
  # $token_ClouDNS = 'xxxxxx'
  # $token_Combell = 'xxxxxx'
  # $token_Constellix = 'xxxxxx'
  # $token_CoreNetworks = 'xxxxxx'
  # $token_DeSEC = 'xxxxxx'
  # $token_DMEasy = 'xxxxxx'
  # $token_DNSimple = 'xxxxxx'
  # $token_DNSPod = 'xxxxxx'
  # $token_DOcean = 'xxxxxx'
  # $token_DomainOffensive = 'xxxxxx'
  # $token_Domeneshop = 'xxxxxx'
  # $token_Dreamhost = 'xxxxxx'
  # $token_DuckDNS = 'xxxxxx'
  # $token_Dynu = 'xxxxxx'
  # $token_EasyDNS = 'xxxxxx'
  # $token_Easyname = 'xxxxxx'
  # $token_FreeDNS = 'xxxxxx'
  # $token_Gandi = 'xxxxxx'
  # $token_GCloud = 'xxxxxx'
  # $token_Google = 'xxxxxx'
  # $token_GoDaddy = 'xxxxxx'
  # $token_Hetzner = 'xxxxxx'
  # $token_HostingDe = 'xxxxxx'
  # $token_HurricaneElectric = 'xxxxxx'
  # $token_IBMSoftLayer = 'xxxxxx'
  # $token_Infoblox = 'xxxxxx'
  # $token_Infomaniak = 'xxxxxx'
  # $token_IONOS = 'xxxxxx'
  # $token_ISPConfig = 'xxxxxx'
  # $token_LeaseWeb = 'xxxxxx'
  # $token_Linode = 'xxxxxx'
  # $token_Loopia = 'xxxxxx'
  # $token_LuaDns = 'xxxxxx'
  # $token_Manual = 'xxxxxx'
  # $token_Namecheap = 'xxxxxx'
  # $token_NameCom = 'xxxxxx'
  # $token_NS1 = 'xxxxxx'
  # $token_OVH = 'xxxxxx'
  # $token_PointDNS = 'xxxxxx'
  # $token_Porkbun = 'xxxxxx'
  # $token_PortsManagement = 'xxxxxx'
  # $token_Rackspace = 'xxxxxx'
  # $token_Regru = 'xxxxxx'
  # $token_RFC2136 = 'xxxxxx'
  # $token_Route53 = 'xxxxxx'
  # $token_Selectel = 'xxxxxx'
  # $token_SimpleDNSPlus = 'xxxxxx'
  # $token_Simply = 'xxxxxx'
  # $token_SSHProxy = 'xxxxxx'
  # $token_TotalUptime = 'xxxxxx'
  # $token_UKFast = 'xxxxxx'
  # $token_Windows = 'xxxxxx'
  # $token_Yandex = 'xxxxxx'
  # $token_Zilore = 'xxxxxx'
  # $token_Zonomi = 'xxxxxx'

  #your exchange server/s
  $exchangeServers = "exchange01.example.com","exchange02.example.com"
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
# LE_PROD = LetsEncrypt productive
# LE_STAGE = LetsEncrypt testing
# BUYPASS_PROD = Buypass productive
# BUYPASS_TEST = Buypass testing
# ZEROSSL_PROD = ZeroSSL productive
#or type your own acme server url
Set-PAServer LE_PROD

#if first time using PoSh-ACME, create ACME-Account with mail, accept terms of use and set private key to rsa
if(!((Get-PAAccount -List).id.count) -ge 1){
    New-PAAccount -AcceptTOS -Contact $certMail -KeyLength 4096
}

#if no cert was yet requested
if((Get-PACertificate $certNames[0]).Subject.count -eq 0){
    #set Plugin Args for HetznerDNS
    $tokens=Get-Variable -Name "token_*"
    foreach($token in $tokens){
      if($null -ne $token.Value){
        $pluginName = $($token.Name.Replace('token_',''))
        $tokenName = $pluginName+"Token"
        $pArgs += @{
          "$tokenName" = ConvertTo-SecureString -AsPlainText -Force -String $($token.value)
        }
      } else {
        Write-Error "No DNS-Plugin specified"
        exit
      }
    }
    if ($null -ne $tokens){
      $command += "-Plugin $plugins -PluginArgs `$pArgs"
    } else {
      Write-Error "No DNS-Plugin specified"
      exit
    }
    #request new certificate
    New-PACertificate $certNames $command -AcceptTOS -Contact $certMail
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

$RxCs=Get-ReceiveConnector
$TxCs=Get-SendConnector

foreach($RxC in $RxCs){
  if($null -ne $RxC.TlsCertificateName){
    $ExCert = Get-ExchangeCertificate -Thumbprint "$($cert.Thumbprint)"
    $TLSCertificateName = "<i>$($ExCert.Issuer)<s>$($ExCert.Subject)"
    Set-ReceiveConnector -Identity $($RxC.Identity) -TlsCertificateName $TLSCertificateName
  }
}

foreach($TxC in $TxCs){
  if($null -ne $TxC.TlsCertificateName){
    $ExCert = Get-ExchangeCertificate -Thumbprint "$($cert.Thumbprint)"
    $TLSCertificateName = "<i>$($ExCert.Issuer)<s>$($ExCert.Subject)"
    Set-ReceiveConnector -Identity $($TxC.Identity) -TlsCertificateName $TLSCertificateName
  }
}