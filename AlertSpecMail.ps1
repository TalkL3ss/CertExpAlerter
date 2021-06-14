Import-Module 'D:\CertExpAlerter\PowerShell\PSPKI-v3.2.7.0\PSPKI.psm1'
$pkiServer = $env:COMPUTERNAME+"."+$env:USERDNSDOMAIN
$pkiMailFrom = $env:COMPUTERNAME+"@"+$env:USERDNSDOMAIN
$mailServer = 'smtp.internal.local'
if (!(Test-Path C:\temp)) { mkdir c:\temp\}
$allIssued = (Get-CertificationAuthority -ComputerName $pkiServer | Get-IssuedRequest -Property CommonName,EMail | ? {$_.Email -and $_.NotAfter -le (Get-Date).AddDays(45) -and $_.NotAfter -ge (Get-Date)} | sort CommonName -Unique | sort NotAfter | select Request.RequesterName,CommonName,CertificateTemplate,Email,NotAfter)
$eMaillist = ($allIssued | Group-Object Email)

foreach ($sMail in $eMaillist) {
    $_tmp = ($allIssued | ? { $_.Email -like $sMail.Name} | ConvertTo-Html | tee C:\temp\cert.html -Append)
    $_tmp = gc C:\temp\cert.html | Out-String
    Send-MailMessage -SmtpServer $mailServer -BodyAsHtml -Body $_tmp  -From $pkiMailFrom -To $sMail.Name -Subject "Certificates that will expire in 45 days" 
    rm -Force C:\temp\cert.html
}

$allIssued = ""
$allIssued = (Get-CertificationAuthority -ComputerName $pkiServer | Get-IssuedRequest -Property CommonName | ? {$_.CertificateTemplate -like "Specific - Template" -and $_.NotAfter -le (Get-Date).AddDays(90) -and $_.NotAfter -ge (Get-Date)} | sort CommonName -Unique | sort NotAfter | select Request.RequesterName,CommonName,CertificateTemplate,NotAfter)
if ($allIssued.Count -ge 1) {
    $_tmp = ($allIssued | ConvertTo-Html | tee C:\temp\MFcert.html -Append)
    $_tmp = gc C:\temp\cert.html | Out-String
    Send-MailMessage -SmtpServer $mailServer -BodyAsHtml -Body $_tmp  -From $pkiMailFrom -To Security@Mail.co.il -Subject "Certificates that will expire in 90 days - From the MainFrame Certificate Template" 
    rm -Force C:\temp\cert.html
}
