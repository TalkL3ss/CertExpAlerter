@echo off 
echo =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
echo 	pkiCerts Automatic Parameters
echo 		Version: 1.1
echo 	      Ohad Halali (2017)
echo =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
rem parametrs explian
rem EXPDays maximum days for alerts
rem LogFile Logfile full path location
rem RCPT send mail to who 

set EXPDays=45
set LogFile=D:\CertExpAlerter\CertExp.log
rem for tests only:  set RCPT="\"mymail@internal.local\""
set RCPT="\"Group1@internal.local\", \"Group2@internal.local\", \"mail@internal.local\", \"Last@internal.local\""
set smtpServer=internal.local
set prvtKey=D:\CertExpAlerter\prvtFile.log


rem -=-=-=-=-=- Start automatic parameters do not change -=-=-=-=-=-
set theHost=%ComputerName%.%USERDNSDOMAIN%
set Subject="Certificate Experation In %EXPDays% Days Report - %theHost%"
set MailFrom=%ComputerName%@internal.local
for /f "tokens=3" %%x in ('reg query HKLM\SYSTEM\CurrentControlSet\Services\CertSvc\Configuration\ /v Active') do (set CAServ="%theHost%\%%x")
rem -=-=-=-=-=- END automatic parameters do not change -=-=-=-=-=-

echo Parameters Summery
echo PKI Server: %CAServ%
echo Maximum days to alert: %EXPDays%
echo Log file location: %LogFile%
echo RCPT To: %RCPT%
echo Mail Server: %smtpServer%
echo Mail Subject: %Subject%
echo Mail From: %MailFrom%

GOTO PrivateKeys

:start
echo:Offline Root CRLs will expire at ^<b^>>> %LogFile%
certutil -dump C:\Windows\System32\certsrv\CertEnroll\isracard-Root-CA.crl | find /i "NextUpdate:" >> %LogFile%
echo:^</b^>^<br^>>> %LogFile%
echo:---------------------------------------^<br^>^<br^> >> %LogFile%

rem FOR /L %%i IN (1,1,%EXPDays%) DO D:\CertExpAlerter\CertExpAlerter.exe -c %CAServ% -d %%i >>%LogFile%
SETLOCAL ENABLEDELAYEDEXPANSION
FOR /L %%i IN (1,1,%EXPDays%) DO (
	for /f "tokens=1,2,3 delims=:" %%f in ('D:\CertExpAlerter\CertExpAlerter.exe -c %CAServ% -d %%i') do (
		set _tmp=!_tmp!%%g
		set _tmp=!_tmp:~-4!
		if "!_tmp!" == "days" ( 
			echo %%f^:%%g ^<br^>  >>%LogFile% 
		) else ( 
			echo %%f^:%%g  >>%LogFile%
		)
	) 
)
powershell -ExecutionPolicy bypass -Command "& {$Body = Get-Content -Path %LogFile% |Out-String ; Send-MailMessage -SmtpServer %smtpServer% -to %RCPT% -From %MailFrom% -Body $Body -Subject \"%Subject%\" -BodyAsHtml}"
GOTO END

:PrivateKeys
echo:Files with high risk (maybe private key is stored)^<br^> > %prvtKey%
dir /s /b c:\*.pfx | find /V /i "\windows" >> %prvtKey%
echo:^<br^> >> %prvtKey%
dir /s /b d:\*.pfx | find /V /i "\windows" >> %prvtKey%
echo:^<br^> >> %LogFile%
echo:Files with high risk (Please delete them)^<br^> >> %prvtKey%
echo:----------------------------------------^<br^> >>  %prvtKey%

rem ****Internal use****
Set /a Lines=0
For /f %%j in ('Find "" /v /c ^< %prvtKey%') Do Set /a Lines=%%j
rem ****END*****
if /i %Lines% GEQ 5 (
type  %prvtKey% > %LogFile% 
) else (
echo:> %LogFile%
)
GOTO start

:END
del /q /f D:\CertExpAlerter\*.log
rem if not exist another internal ca remove the line below
%SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe "D:\CertExpAlerter\PowerShell\CertificateAlerter.ps1" -bypass
D:\CertExpAlerter\PsExec.exe -accepteula  \\additinal.internal.ca.local C:\CertExpAlerter\CertExpAlerter.bat
verify
exit
