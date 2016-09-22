Rem ###################
Rem Usage: SPPwdChange.Bat username password
Rem ###################

::modify apppool identity
cd %windir%\System32\inetsrv
appcmd.exe set apppool "SharePoint - 80" /processmodel.userName:%1 /processmodel.password:%2
appcmd.exe set apppool "SharePoint Central Administration v4" /processmodel.userName:$1 /processmodel.password:%2

::modify sharepoint credential
cd "c:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\BIN"
stsadm -o updatefarmcredentials -userlogin %1 -password %2