Attribute VB_Name = "EASendmail"
Const ConnectNormal = 0
Const ConnectSSLAuto = 1
Const ConnectSTARTTLS = 2
Const ConnectDirectSSL = 3
Const ConnectTryTLS = 4

Public Sub EASendMail()

    Dim oSmtp As New EASendMailObjLib.Mail
    oSmtp.LicenseCode = "TryIt"

    ' Set your Gmail email address
    oSmtp.FromAddr = "piki.sendmail@gmail.com"

    ' Add recipient email address
    oSmtp.AddRecipientEx "alisoncf@gmail.com", 0

    ' Set email subject
    oSmtp.subject = "test email from gmail account"

    ' Set email body
    oSmtp.BodyText = "this is a test email sent from VB 6.0 project with gmail"

    ' Gmail SMTP server address
    oSmtp.ServerAddr = "smtp.gmail.com"

    ' If you want to use direct SSL 465 port,
    ' Please add this line, otherwise TLS will be used.
    ' oSmtp.ServerPort = 465

    ' set 25 or 587 port
    oSmtp.ServerPort = 587

    ' detect SSL/TLS automatically
    oSmtp.ConnectType = ConnectSSLAuto

    ' Gmail user authentication should use your
    ' Gmail email address as the user name.
    ' For example: your email is "gmailid@gmail.com", then the user should be "gmailid@gmail.com"
    oSmtp.UserName = "piki.sendmail@gmail.com"

    ' Create app password in Google account
    ' https://support.google.com/accounts/answer/185833?hl=en
    oSmtp.Password = "Anapolis1"

    MsgBox "start to send email ..."

    If oSmtp.SendMail() = 0 Then
        MsgBox "email was sent successfully!"
    Else
        MsgBox "failed to send email with the following error:" & oSmtp.GetLastErrDescription()
    End If

End Sub
