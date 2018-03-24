# # # # # # # # # # # # # # # # # # # # #
# # # # # # # CONFIGURATION # # # # # # #
# # # # # # # # # # # # # # # # # # # # #
$minuteschanged = 60
$distinguishedname = "CN=Users,DC=domain,DC=com"
$emailSmtpServer = "smtp.domain.com"
$emailSmtpServerPort = "587"
$emailSmtpUser = "smtpusername"
$emailSmtpPass = "smtppassword"
$emailFrom = "Your Name <foo@bar.com>"
$emailSubject = "Your password has been reset"
$emailBodyPath = "C:\YourPathHere\email.txt"
# # # # # # # # # # # # # # # # # # # # #


# Determine when $minuteschanged was in the past
$timeminus1=(get-date).addminutes("-$minuteschanged")

$users = get-aduser -searchbase $distinguishedname -filter * -properties name,mail,sAMAccountName,pwdLastSet | Where {[datetime]::fromFileTime($_.pwdLastSet) -ge $timeminus1} | ForEach-Object {
    
    # Send the email to the email address listed for the user in Active Directory
    $to = $_.mail

    # Set the of the user as their Display Name in Active Directory
    $name = $_.name

    # Convert pwdLastSet from a timestamp to a read-able format
    $LastSet = [datetime]::fromFileTime($_.pwdLastSet)

    # Start a new mail message and set the values to send it
    $emailMessage = New-Object System.Net.Mail.MailMessage
    $emailMessage.From = $emailFrom
    $emailMessage.To.Add( $to )
    $emailMessage.Subject = $emailSubject
    $emailMessage.IsBodyHtml = $true
    $emailMessage.Body = Get-content $emailBodyPath
    $emailMessage.Body = invoke-expression $emailMessage.Body
    
    # Connect to the SMTP Server
    $SMTPClient = New-Object System.Net.Mail.SmtpClient( $emailSmtpServer , $emailSmtpServerPort )
    $SMTPClient.EnableSsl = $true
    $SMTPClient.Credentials = New-Object System.Net.NetworkCredential( $emailSmtpUser , $emailSmtpPass );
    
    # Send the email
    $SMTPClient.Send( $emailMessage )
}

Write-Output "The script has completed"