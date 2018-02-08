function ConvertTo-Alias
{
    param
    (
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,Position=0)]
        [string[]]$InputObject
    )
    begin
    {
        $stringBuilder = [System.Text.StringBuilder]::new()
    }
    process
    {
        foreach ($string in $InputObject)
        {
            $string = $string.Trim('.')
            [void]$stringBuilder.Clear()
            [char]$prev = "`0"
            foreach ($char in [char[]]$string.Normalize([System.Text.NormalizationForm]::FormD))
            {
                if ([System.Globalization.CharUnicodeInfo]::GetUnicodeCategory($char) -ne [System.Globalization.UnicodeCategory]::NonSpacingMark)
                {
                    if ($char -eq 'ø' -or $char -eq 'Ø')
                    {
                        $char = 'o'
                    }
                    elseif ($char -eq 'æ' -or $char -eq 'Æ')
                    {
                        $char = 'a'
                    }
                    elseif ($char -eq ' ')
                    {
                        $char = '_'
                    }
                    if (($char -eq '_' -or $char -eq '-') -and $char -eq $prev)
                    {
                        continue
                    }
                    if ($char -eq '_' -and $prev -eq '-')
                    {
                        $prev = '_'
                        continue
                    }
                    if ($char -eq '-' -and $prev -eq '_')
                    {
                        $prev = '-'
                        continue
                    }
                    if ($char -match '[A-Za-z0-9!#$%&''*+/=?^_`{|}~-]')
                    {
                        [void]$stringBuilder.Append($char)
                        $prev = $char
                    }
                }
            }
            Write-Output -InputObject ($stringBuilder.ToString().Trim(@('-','_')))
        }
    }
}
