# PSSigntools
PSSigntools are PowerShells scripts for signing exe/dll with Code signing certificate.

## sign.bat / sign.ps1
### 
` sign <file1> [<file2> ...]`
### Overview
Sign argument files with Code signing certificate.<br>
Following files are available.
* Executables(.exe .dll)
* PowerShell scripts(.ps1)

If a file already has a valid signature, the file is not signed.

sign.bat is a batch file calling sign.ps1.

## updatevsto.bat / updatevsto.ps1
### Usage
`  updatevsto.bat <file.vsto>`
### Overview

Sign Microsoft Office add-in DLL file and related files like manifest file (\*.manifest) and VSTO installer file (\*.vsto).

If the DLL file already has a valid signature, the file is not signed.

updatevsto.bat is a batch file calling updatevsto.ps1.

## NOTICE

Sign/updatevsto works on the assumption that only one valid Code signing certificate is registered on the computer.
When you renew the certificate, you will need uninstall old certificate.

Timestamp server uses Globalsign(http://timestamp.globalsign.com/scripts/timstamp.dll).
It is embedded in script, so if you want to change timestamp server, you have to change script.
