function Set-GroupData
{
    param
    (
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [int]
        $DatabaseId,

        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
        [object]
        $ObjectGuid,

        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
        [string]
        $GeneratedAlias,

        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
        [string]
        $GeneratedDisplayName,

        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
        [object]
        $CreatedOn,

        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
        [string]
        $ErrorText
    )
    $params = @{
        Name = 'spO365SetGroupData'
        Type = 'NonQuery'
        Parameter = @{id=$DatabaseId}
    }
    if ($ObjectGuid)
    {
        $params.Parameter.guid = $ObjectGuid
    }
    if ($GeneratedAlias)
    {
        $params.Parameter.alias = $GeneratedAlias
    }
    if ($GeneratedDisplayName)
    {
        $params.Parameter.displayName = $GeneratedDisplayName
    }
    if ($CreatedOn)
    {
        $params.Parameter.created = $CreatedOn
    }
    if ($ErrorText)
    {
        $params.Parameter.error = $ErrorText
    }
    [void](Invoke-StoredProcedure @params)
}

function Get-GroupData
{
    $groups = Invoke-StoredProcedure -Name 'spO365GetGroupData' -Type 'Reader'
    foreach ($group in $groups)
    {
        $output = [pscustomobject]@{
            DatabaseId = $group.id
            RequestedName = $group.groupName
            RequestedSuffix = $group.groupSuffix
            RequestedOwner = $group.groupOwner
            ObjectGuid = $group.groupGuid
            NamingScheme = $group.namingScheme
            GeneratedAlias = $group.generatedAlias
            GeneratedDisplayName = $group.generatedDisplayName
            CreatedOn = $group.created
            DuplicateDisplayName = $false
            ErrorText = $group.error
        }
        Write-Output -InputObject $output
    }
}

function Invoke-StoredProcedure
{
    param
    (
        [Parameter(Mandatory=$true)]
        [string]
        $Name,
        [Parameter(Mandatory=$true)]
        [string]
        [ValidateSet('Reader','Scalar','NonQuery')]
        $Type,
        [hashtable]
        $Parameter
    )
    $params = @{
        CommandType = 'StoredProcedure'
        CommandText = "dbo.$Name"
    }
    if ($Parameter)
    {
        $params.Parameter = $Parameter
    }
    $command = New-SqlCommand @params
    switch ($Type)
    {
        'Reader' {
            $reader = $command.ExecuteReader()
            while ($reader.Read())
            {
                $props = @{}
                for ($i = 0; $i -lt $reader.FieldCount; $i++)
                {
                    $value = $reader.GetValue($i)
                    if ($value -is [System.DBNull])
                    {
                        $value = $null
                    }
                    $props.Add($reader.GetName($i), $value)
                }
                Write-Output -InputObject ([pscustomobject]$props)
            }
            $reader.Dispose()
            break
        }
        'Scalar' {
            Write-Output -InputObject $command.ExecuteScalar()
            break
        }
        'NonQuery' {
            Write-Output -InputObject ([int]$command.ExecuteNonQuery())
            break
        }
    }
    $command.Dispose()
}

function New-SqlCommand
{
    param
    (
        [System.Data.CommandType]
        $CommandType,
        [string]
        $CommandText,
        [hashtable]
        $Parameter
    )

    $command = [System.Data.SqlClient.SqlCommand]@{
        Connection = Get-SqlConnection
        CommandType = $CommandType
        CommandText = $CommandText
    }
    if ($Parameter -ne $null)
    {
        foreach ($item in $Parameter.GetEnumerator())
        {
            [void]$command.Parameters.AddWithValue($item.Name, $item.Value)
        }
    }
    Write-Output -InputObject $command
}

function Get-SqlConnection
{
    if (-not $Script:SqlConnection)
    {
        $connectionString = "Server=$($Script:Config.SqlServer);Database=$($Script:Config.SqlDatabase);Trusted_Connection=True;"
        $Script:SqlConnection = [System.Data.SqlClient.SqlConnection]($connectionString)
        $Script:SqlConnection.Open()
    }
    Write-Output -InputObject $SqlConnection
}
