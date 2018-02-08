# Make all errors terminating errors
$ErrorActionPreference = 'Stop'

. "$PSScriptRoot\Config.ps1"
. "$PSScriptRoot\SqlBackend.ps1"
. "$PSScriptRoot\Utils.ps1"

Get-PSSession -Name 'ExchangeOnline' -ErrorAction SilentlyContinue | Remove-PSSession
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
}
Import-PSSession @params | Out-Null
