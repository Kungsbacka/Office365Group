$Script:Config = @{
    O365User = '<office 365 user that can manage groups>'
    O365Password = '<encrypted password>'
    SqlServer = '<server instance>'
    SqlDatabase = '<database>'
    SmtpServer = '<smtp server without authentication>'
    SmtpFrom = 'smtp sender'
    SmtpSubject = 'New Office 365 group ready'
    SmtpBody = "Hi,`r`n`r`nYour new group ""{0}"" is now available in Office 365. `r`n`r`nEnjoy!"
}
