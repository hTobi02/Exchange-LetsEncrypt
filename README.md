# New-LEExchangeCertificate
Using this script, a LetsEncrypt certificate is automatically requested and imported into an Exchange server and included for the IMAP,POP,SMTP and IIS services. 

## Dependencies
You need the Powershell module "PoSh-ACME". It will be imported at the beginning of the script. If it is not installed, the script tries to install it in the user context, because no administration rights are needed for this. <br>
You can get more information about PoSH-ACME [HERE](https://poshac.me/docs/v4/#setup)

For questions, suggestions or to report bugs please open an [Issue](https://github.com/hTobi02/Exchange-LetsEncrypt/issues/new/choose) or a [PR](https://github.com/hTobi02/Exchange-LetsEncrypt/compare). 