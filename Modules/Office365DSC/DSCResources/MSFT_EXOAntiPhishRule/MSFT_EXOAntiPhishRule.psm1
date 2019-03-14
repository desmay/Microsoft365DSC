function Get-TargetResource
{
    [CmdletBinding()]
    [OutputType([System.Collections.Hashtable])]
    param
    (
        [Parameter(Mandatory = $true)]
        [System.String]
        $Identity,

        [Parameter(Mandatory = $true)]
        [System.String]
        $AntiPhishPolicy,

        [Parameter()]
        [System.String]
        $Comments,

        [Parameter()]
        [System.Boolean]
        $Enabled = $true,

        [Parameter()]
        [ValidateSet('Present', 'Absent')]
        [System.String]
        $Ensure = 'Present',

        [Parameter()]
        [System.String[]]
        $ExceptIfRecipientDomainIs = @(),

        [Parameter()]
        [System.String[]]
        $ExceptIfSentTo = @(),

        [Parameter()]
        [System.String[]]
        $ExceptIfSentToMemberOf = @(),

        [Parameter()]
        [uint32]
        $Priority,

        [Parameter()]
        [System.String[]]
        $RecipientDomainIs = @(),

        [Parameter()]
        [System.String[]]
        $SentTo = @(),

        [Parameter()]
        [System.String[]]
        $SentToMemberOf = @(),

        [Parameter(Mandatory = $true)]
        [System.Management.Automation.PSCredential]
        $GlobalAdminAccount
    )
    Write-Verbose "Get-TargetResource will attempt to retrieve AntiPhishRule $($Identity)"
    Write-Verbose "Calling Connect-ExchangeOnline function:"
    Connect-ExchangeOnline -GlobalAdminAccount $GlobalAdminAccount
    Write-Verbose "Global ExchangeOnlineSession status:"
    Write-Verbose "$( Get-PSSession -ErrorAction SilentlyContinue | Where-Object Name -eq 'ExchangeOnline' | Out-String)"
    try
    {
        $AntiPhishRules = Get-AntiPhishRule
    }
    catch
    {
        Close-SessionsAndReturnError -ExceptionMessage $_.Exception
    }
    $AntiPhishRule = $AntiPhishRules | Where-Object Identity -eq $Identity
    if (-NOT $AntiPhishRule)
    {
        Write-Verbose "AntiPhishRule $($Identity) does not exist."
        $result = @{
            Ensure             = 'Absent'
            GlobalAdminAccount = $GlobalAdminAccount
            Identity           = $Identity
        }
        return $result
    }
    else
    {
        $result = @{
            Ensure = 'Present'
        }
        foreach ($KeyName in ($PSBoundParameters.Keys | Where-Object {$_ -ne 'Ensure'}) )
        {
            if ($null -ne $AntiPhishRule.$KeyName)
            {
                $result += @{
                    $KeyName = $AntiPhishRule.$KeyName
                }
            }
            else
            {
                $result += @{
                    $KeyName = $PSBoundParameters[$KeyName]
                }
            }

        }
        if ('Enabled' -eq $AntiPhishRule.State)
        {
            # Accounts for Get-AntiPhishRule returning 'State' instead of 'Enabled' used by New/Set
            $result.Enabled = $true
        }
        else
        {
            $result.Enabled = $false
        }

        Write-Verbose "Found AntiPhishRule $($Identity)"
        Write-Verbose "Get-TargetResource Result: `n $($result | Out-String)"
        return $result
    }
}

function Set-TargetResource
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [System.String]
        $Identity,

        [Parameter(Mandatory = $true)]
        [System.String]
        $AntiPhishPolicy,

        [Parameter()]
        [System.String]
        $Comments,

        [Parameter()]
        [System.Boolean]
        $Enabled = $true,

        [Parameter()]
        [ValidateSet('Present', 'Absent')]
        [System.String]
        $Ensure = 'Present',

        [Parameter()]
        [System.String[]]
        $ExceptIfRecipientDomainIs = @(),

        [Parameter()]
        [System.String[]]
        $ExceptIfSentTo = @(),

        [Parameter()]
        [System.String[]]
        $ExceptIfSentToMemberOf = @(),

        [Parameter()]
        [uint32]
        $Priority,

        [Parameter()]
        [System.String[]]
        $RecipientDomainIs = @(),

        [Parameter()]
        [System.String[]]
        $SentTo = @(),

        [Parameter()]
        [System.String[]]
        $SentToMemberOf = @(),

        [Parameter(Mandatory = $true)]
        [System.Management.Automation.PSCredential]
        $GlobalAdminAccount
    )
    Write-Verbose 'Entering Set-TargetResource'
    Write-Verbose 'Retrieving information about AntiPhishRule configuration'
    Write-Verbose "Calling Connect-ExchangeOnline function:"
    Connect-ExchangeOnline -GlobalAdminAccount $GlobalAdminAccount
    Write-Verbose "Global ExchangeOnlineSession status:"
    Write-Verbose "$( Get-PSSession -ErrorAction SilentlyContinue | Where-Object Name -eq 'ExchangeOnline' | Out-String)"
    try
    {
        $AntiPhishRules = Get-AntiPhishRule
    }
    catch
    {
        Close-SessionsAndReturnError -ExceptionMessage $_.Exception
    }

    $AntiPhishRule = $AntiPhishRules | Where-Object Identity -eq $Identity

    if ( ('Present' -eq $Ensure ) -and (-NOT $AntiPhishRule) )
    {
        try
        {
            NewAntiPhishRule -NewAntiPhishRuleParams $PSBoundParameters
        }
        catch
        {
            Close-SessionsAndReturnError -ExceptionMessage $_.Exception
        }
    }

    if ( ('Present' -eq $Ensure ) -and ($AntiPhishRule) )
    {
        try
        {
            if ($PSBoundParameters.Enabled -and ('Disabled' -eq $AntiPhishRule.State) )
            {
                # New-AntiPhishRule has the Enabled parameter, Set-AntiPhishRule does not.
                # There doesn't appear to be any way to change the Enabled state of a rule once created.
                Write-Verbose "Removing AntiPhishRule $($Identity) in order to change Enabled state."
                Remove-AntiPhishRule -Identity $Identity -Confirm:$false
                NewAntiPhishRule -NewAntiPhishRuleParams $PSBoundParameters
            }
            else
            {
                SetAntiPhishRule -SetAntiPhishRuleParams $PSBoundParameters
            }
        }
        catch
        {
            Close-SessionsAndReturnError -ExceptionMessage $_.Exception
        }
    }

    if ( ('Absent' -eq $Ensure ) -and ($AntiPhishRule) )
    {
        Write-Verbose "Removing AntiPhishRule $($Identity) "
        try
        {
            Remove-AntiPhishRule -Identity $Identity -Confirm:$false
        }
        catch
        {
            Close-SessionsAndReturnError -ExceptionMessage $_.Exception
        }
    }

    Write-Verbose "Closing Remote PowerShell Sessions"
    $ClosedPSSessions = (Get-PSSession | Remove-PSSession)
    Write-Verbose "Global ExchangeOnlineSession status: `n"
    Write-Verbose "$( Get-PSSession -ErrorAction SilentlyContinue | Where-Object Name -eq 'ExchangeOnline' | Out-String)"
}

function Test-TargetResource
{
    [CmdletBinding()]
    [OutputType([System.Boolean])]
    param
    (
        [Parameter(Mandatory = $true)]
        [System.String]
        $Identity,

        [Parameter(Mandatory = $true)]
        [System.String]
        $AntiPhishPolicy,

        [Parameter()]
        [System.String]
        $Comments,

        [Parameter()]
        [System.Boolean]
        $Enabled = $true,

        [Parameter()]
        [ValidateSet('Present', 'Absent')]
        [System.String]
        $Ensure = 'Present',

        [Parameter()]
        [System.String[]]
        $ExceptIfRecipientDomainIs = @(),

        [Parameter()]
        [System.String[]]
        $ExceptIfSentTo = @(),

        [Parameter()]
        [System.String[]]
        $ExceptIfSentToMemberOf = @(),

        [Parameter()]
        [uint32]
        $Priority,

        [Parameter()]
        [System.String[]]
        $RecipientDomainIs = @(),

        [Parameter()]
        [System.String[]]
        $SentTo = @(),

        [Parameter()]
        [System.String[]]
        $SentToMemberOf = @(),

        [Parameter(Mandatory = $true)]
        [System.Management.Automation.PSCredential]
        $GlobalAdminAccount
    )
    Write-Verbose -Message "Testing AntiPhishRule for $($Identity)"
    $CurrentValues = Get-TargetResource @PSBoundParameters
    $AntiPhishRuleTestParams = $PSBoundParameters
    $AntiPhishRuleTestParams.Remove("GlobalAdminAccount") | out-null
    $TestResult = Test-Office365DSCParameterState -CurrentValues $CurrentValues `
        -DesiredValues $PSBoundParameters `
        -ValuesToCheck $AntiPhishRuleTestParams.Keys
    if ($TestResult)
    {
        Write-Verbose "Closing Remote PowerShell Sessions"
        $ClosedPSSessions = (Get-PSSession | Remove-PSSession)
        Write-Verbose "Global ExchangeOnlineSession status: `n"
        Write-Verbose "$( Get-PSSession -ErrorAction SilentlyContinue | Where-Object Name -eq 'ExchangeOnline' | Out-String)"
    }

    return $TestResult
}

function Export-TargetResource
{
    [CmdletBinding()]
    [OutputType([System.String])]
    param
    (
        [Parameter(Mandatory = $true)]
        [System.String]
        $Identity,

        [Parameter(Mandatory = $true)]
        [System.String]
        $AntiPhishPolicy,

        [Parameter(Mandatory = $true)]
        [System.Management.Automation.PSCredential]
        $GlobalAdminAccount
    )
    $result = Get-TargetResource @PSBoundParameters
    Write-Verbose "Closing Remote PowerShell Sessions"
    $ClosedPSSessions = (Get-PSSession | Remove-PSSession)
    Write-Verbose "Global ExchangeOnlineSession status: `n"
    Write-Verbose "$( Get-PSSession -ErrorAction SilentlyContinue | Where-Object Name -eq 'ExchangeOnline' | Out-String)"
    $result.GlobalAdminAccount = Resolve-Credentials -UserName $GlobalAdminAccount.UserName
    $content = "        EXOAntiPhishRule " + (New-GUID).ToString() + "`r`n"
    $content += "        {`r`n"
    $currentDSCBlock = Get-DSCBlock -Params $result -ModulePath $PSScriptRoot
    $content += Convert-DSCStringParamToVariable -DSCBlock $currentDSCBlock -ParameterName "GlobalAdminAccount"
    $content += "        }`r`n"
    return $content
}

Export-ModuleMember -Function *-TargetResource