# Make all errors terminating errors
$ErrorActionPreference = 'Stop'

. "$PSScriptRoot\Config.ps1"
. "$PSScriptRoot\SqlBackend.ps1"
. "$PSScriptRoot\Utils.ps1"

$groupData = Get-GroupData

If ($groupData -eq $null)
{
    exit
}
  
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
            'Get-UnifiedGroup'
            'New-UnifiedGroup'
            'Set-UnifiedGroup'
            'Add-UnifiedGroupLinks'
            'Remove-UnifiedGroupLinks'
        )
    }
    Import-PSSession @params | Out-Null
}

# Generate names and check for duplicates
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
            break
        }
        default
        {
            $item.ErrorText = "Unknown naming scheme ""$($item.NamingScheme)"""
            $item | Set-GroupData
            continue
        }
    }
    try
    {
        if (Get-UnifiedGroup -Filter "DisplayName -eq '$($item.GeneratedDisplayName)'")
        {
            $item.ErrorText = "A group with the display name ""$($item.GeneratedDisplayName)"" already exists"
            $item.DuplicateDisplayName = $true
        }
    }
    catch
    {
        $item.ErrorText = $_.Exception
    }
    $item | Set-GroupData
}

# Create new groups
foreach ($item in $groupData)
{
    if ($item.ErrorText -ne $null)
    {
        continue
    }
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
foreach ($item in $groupData)
{
    $msg = [System.Net.Mail.MailMessage]@{
        BodyEncoding = [System.Text.Encoding]::UTF8
        SubjectEncoding = [System.Text.Encoding]::UTF8
        From = $Script:Config.SmtpFrom
    }
    $msg.To.Add($item.RequestedOwner)
    if ($item.ErrorText -eq $null -and -not $item.DuplicateDisplayName)
    {
        $msg.Subject = $Script:Config.SuccessMail.Subject
        $msg.Body = $Script:Config.SuccessMail.Body -f $item.GeneratedDisplayName
    }
    elseif ($item.DuplicateDisplayName)
    {
        $msg.Subject = $Script:Config.DuplicateMail.Subject
        $msg.Body = $Script:Config.DuplicateMail.Body -f $item.GeneratedDisplayName
    }
    else
    {
        $msg.Subject = $Script:Config.ErrorMail.Subject
        $msg.Body = $Script:Config.ErrorMail.Body -f $item.GeneratedDisplayName
    }
    try
    {
        $smtpClient.Send($msg)
    }
    catch
    {
        $item.ErrorText = $_.Exception
    }
    $msg.Dispose()
    $item | Set-GroupData
}
