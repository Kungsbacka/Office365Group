$Script:Config = @{
    O365User = '<office 365 user that can manage groups>'
    O365Password = '<encrypted password>'
    SqlServer = '<server instance>'
    SqlDatabase = '<database>'
    SmtpServer = '<smtp server without authentication>'
    SmtpFrom = 'smtp sender'
    SuccessMail = @{
        Subject = 'New Office 365 group ready'
        Body = "Hi,`r`n`r`nYour new group ""{0}"" is now available in Office 365. `r`n`r`nEnjoy!"
    }
    ErrorMail = @{        Subject = 'Create Office 365 group failed'
        Body = "Hi,`r`n`r`nThere was an error when the group ""{0}"" was created."
    }
    DuplicateMail = @{
        Subject = 'Duplicate Office 365 group'
        Body = "Hi,`r`n`r`nYour new group ""{0}"" has the same name as an existing group in Office 365. Chose another name and try again."
    }

}
