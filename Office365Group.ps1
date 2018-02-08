# Make all errors terminating errors
$ErrorActionPreference = 'Stop'

. "$PSScriptRoot\Config.ps1"
. "$PSScriptRoot\SqlBackend.ps1"
. "$PSScriptRoot\Utils.ps1"

if (-not (Get-PSSession -Name 'ExchangeOnline' -ErrorAction SilentlyContinue))
{
    $params = @{
        Name = 'ExchangeOnline'
        ConfigurationName = 'Microsoft.Exchange'
        ConnectionUri = 'https://outlook.office365.com/powershell-liveid/'
        Authentication = 'Basic'
        AllowRedirection = $true
        Credential = [System.Management.Automation.PSCredential]::new(
            $Script:Config.O365User,
            ($Script:Config.O365Password | ConvertTo-SecureString)
        )
    }
    $session = New-PSSession @params
    $params = @{
        Session = $session
        DisableNameChecking = $true
        CommandName = @(
            'New-UnifiedGroup'
            'Set-UnifiedGroup'
            'Add-UnifiedGroupLinks'
            'Remove-UnifiedGroupLinks'
        )
    }
    Import-PSSession @params | Out-Null
}

$groupData = Get-GroupData -Create
foreach ($item in $groupData)
{
    $aliasSuffix = '_' + ([guid]::NewGuid() -split '-')[0]
    switch ($item.NamingScheme)
    {
        'FG'
        {
            if (-not $item.RequestedSuffix)
            {
                Write-Error -Message "Group suffix is missing. Naming scheme ""$($item.NamingScheme)"" requires a suffix."
                continue
            }
            $today = [datetime]::Today
            $sem = $today.ToString('yy')
            # Cutoff date approx. june 15. This will be wrong on a leap year, but that's ok.
            if ($today.DayOfYear -gt 166)
            {
                $sem = 'HT' + $sem
            }
            else
            {
                $sem = 'VT' + $sem
            }
            $item.GeneratedDisplayName = "$($item.RequestedName) $($item.RequestedSuffix) $sem"
            $item.GeneratedAlias = (ConvertTo-Alias -InputObject $item.GeneratedDisplayName) + $aliasSuffix
            $item | Set-GroupData
            break
        }
        default
        {
            $item.ErrorText = "Unknown naming scheme ""$($item.NamingScheme)"""
            $item | Set-GroupData
            continue
        }
    }
}

foreach ($item in $groupData)
{
    $params = @{
        DisplayName = $item.GeneratedDisplayName
        Alias = $item.GeneratedAlias
        AccessType = 'Private'
        Language = 'sv-SE'
        AutoSubscribeNewMembers = $true
        Members =  @($item.RequestedOwner)
    }
    try
    {
        $unifiedGroup = New-UnifiedGroup @params
        $item.ObjectGuid = $unifiedGroup.Guid
        Add-UnifiedGroupLinks -Identity $unifiedGroup.Identity -Links @($item.RequestedOwner) -LinkType 'Owners'
        Remove-UnifiedGroupLinks -Identity $unifiedGroup.Identity -Links @($Script:Config.O365User) -LinkType 'Owners' -Confirm:$false
        Remove-UnifiedGroupLinks -Identity $unifiedGroup.Identity -Links @($Script:Config.O365User) -LinkType 'Members' -Confirm:$false
        $item.CreatedOn = Get-Date
    }
    catch
    {
        $item.ErrorText = $_.Exception
    }
    $item | Set-GroupData
}

$smtpClient = [System.Net.Mail.SmtpClient]@{
    UseDefaultCredentials = $false
    Host = $Script:Config.SmtpServer
}
$groupData = Get-GroupData -Report
foreach ($item in $groupData)
{
    $msg = [System.Net.Mail.MailMessage]@{
        BodyEncoding = [System.Text.Encoding]::UTF8
        SubjectEncoding = [System.Text.Encoding]::UTF8
        From = $Script:Config.SmtpFrom
        Subject = $Script:Config.SmtpSubject
        Body = $Script:Config.SmtpBody -f $item.GeneratedDisplayName
    }
    $msg.To.Add($item.RequestedOwner)
    try
    {
        $smtpClient.Send($msg)
        $item.ReportedOn = Get-Date
    }
    catch
    {
        $item.ErrorText = $_.Exception
    }
    $msg.Dispose()
    $item | Set-GroupData
}
