function Get-TargetResource
{
    [CmdletBinding()]
    [OutputType([System.Collections.Hashtable])]
    param
    (
        [Parameter(Mandatory = $true)]
        [System.String]
        $Name,

        [Parameter()]
        [System.String]
        $Comment,

        [Parameter()]
        [System.DateTime]
        $EventDateTime,

        [Parameter()]
        [System.String[]]
        $EventTags,

        [Parameter()]
        [System.String[]]
        $EventTypes,

        [Parameter()]
        [Microsoft.Management.Infrastructure.CimInstance[]]
        $AssetId,

        [Parameter()]
        [Microsoft.Management.Infrastructure.CimInstance[]]
        $ExchangeAssetIdQuery,

        [Parameter()]
        [Microsoft.Management.Infrastructure.CimInstance[]]
        $SharePointAssetIdQuery,

        [Parameter()]
        [ValidateSet('Present', 'Absent')]
        [System.String]
        $Ensure = 'Present',

        [Parameter(Mandatory = $true)]
        [System.Management.Automation.PSCredential]
        $GlobalAdminAccount
    )

    Write-Verbose -Message "Getting configuration of Retention Event for $Name"
    #region Telemetry
    $ResourceName = $MyInvocation.MyCommand.ModuleName.Replace("MSFT_", "")
    $data = [System.Collections.Generic.Dictionary[[String], [String]]]::new()
    $data.Add("Resource", $ResourceName)
    $data.Add("Method", $MyInvocation.MyCommand)
    $data.Add("Principal", $GlobalAdminAccount.UserName)
    Add-M365DSCTelemetryEvent -Data $data
    #endregion

    if ($Global:CurrentModeIsExport)
    {
        $ConnectionMode = New-M365DSCConnection -Platform 'SecurityComplianceCenter' `
            -InboundParameters $PSBoundParameters `
            -SkipModuleReload $true
    }
    else
    {
        $ConnectionMode = New-M365DSCConnection -Platform 'SecurityComplianceCenter' `
            -InboundParameters $PSBoundParameters
    }

    $nullReturn = $PSBoundParameters
    $nullReturn.Ensure = 'Absent'
    try
    {
        $EventObject = Get-ComplianceRetentionEvent -Identity $Name `
            -ErrorAction SilentlyContinue

        if ($null -eq $EventObject)
        {
            Write-Verbose -Message "RetentionComplianceEvent $($Name) does not exist."
            return $nullReturn
        }
        else
        {
            foreach ($eventTag in $EventObject.eventTags)
            {
                $complianceTag = Get-ComplianceTag -identity $eventTag
                if ($null -ne $complianceTag)
                {
                    $complianceTags += $complianceTag.Name
                }
            }


            Write-Verbose "Found existing RetentionComplianceEvent $($Name)"
            $result = @{
                Name                   = $EventObject.Name
                Comment                = $EventObject.Comment
                GlobalAdminAccount     = $GlobalAdminAccount
                EventDateTime          = $EventObject.EventDateTime
                EventTags              = $complianceTags
                AssetId                = $EventObject.AssetId
                ExchangeAssetIdQuery   = $EventObject.ExchangeAssetIdQuery
                SharePointAssetIdQuery = $EventObject.SharePointAssetIdQuery
                Ensure                 = 'Present'
            }

            Write-Verbose -Message "Found RetentionComplianceEvent $($Name)"
            Write-Verbose -Message "Get-TargetResource Result: `n $(Convert-M365DscHashtableToString -Hashtable $result)"
            return $result
        }
    }
    catch
    {
        try
        {
            Write-Verbose -Message $_
            $tenantIdValue = ""
            if (-not [System.String]::IsNullOrEmpty($TenantId))
            {
                $tenantIdValue = $TenantId
            }
            elseif ($null -ne $GlobalAdminAccount)
            {
                $tenantIdValue = $GlobalAdminAccount.UserName.Split('@')[1]
            }
            Add-M365DSCEvent -Message $_ -EntryType 'Error' `
                -EventID 1 -Source $($MyInvocation.MyCommand.Source) `
                -TenantId $tenantIdValue
        }
        catch
        {
            Write-Verbose -Message $_
        }
        return $nullReturn
    }
}

function Set-TargetResource
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [System.String]
        $Name,

        [Parameter()]
        [System.String]
        $Comment,

        [Parameter()]
        [System.DateTime]
        $EventDateTime,

        [Parameter()]
        [System.String[]]
        $EventTags,

        [Parameter()]
        [System.String[]]
        $EventTypes,

        [Parameter()]
        [Microsoft.Management.Infrastructure.CimInstance[]]
        $AssetId,

        [Parameter()]
        [Microsoft.Management.Infrastructure.CimInstance[]]
        $ExchangeAssetIdQuery,

        [Parameter()]
        [Microsoft.Management.Infrastructure.CimInstance[]]
        $SharePointAssetIdQuery,

        [Parameter()]
        [ValidateSet('Present', 'Absent')]
        [System.String]
        $Ensure = 'Present',

        [Parameter(Mandatory = $true)]
        [System.Management.Automation.PSCredential]
        $GlobalAdminAccount
    )

    Write-Verbose -Message "Setting configuration of RetentionComplianceEventType for $Name"
    #region Telemetry
    $ResourceName = $MyInvocation.MyCommand.ModuleName.Replace("MSFT_", "")
    $data = [System.Collections.Generic.Dictionary[[String], [String]]]::new()
    $data.Add("Resource", $ResourceName)
    $data.Add("Method", $MyInvocation.MyCommand)
    $data.Add("Principal", $GlobalAdminAccount.UserName)
    Add-M365DSCTelemetryEvent -Data $data
    #endregion

    $ConnectionMode = New-M365DSCConnection -Platform 'SecurityComplianceCenter' `
        -InboundParameters $PSBoundParameters

    $CurrentEventType = Get-TargetResource @PSBoundParameters

    if (('Present' -eq $Ensure) -and ('Absent' -eq $CurrentEventType.Ensure))
    {
        $CreationParams = $PSBoundParameters
        $CreationParams.Remove("GlobalAdminAccount") | Out-Null
        $CreationParams.Remove("Ensure") | Out-Null
        New-ComplianceRetentionEventType @CreationParams
    }
    elseif (('Present' -eq $Ensure) -and ('Present' -eq $CurrentEventType.Ensure))
    {
        $CreationParams = $PSBoundParameters
        $CreationParams.Remove("GlobalAdminAccount") | Out-Null
        $CreationParams.Remove("Ensure") | Out-Null
        $CreationParams.Remove("Name") | Out-Null
        $CreationParams.Add("Identity", $Name)
        Set-ComplianceRetentionEventType @CreationParams
    }
    elseif (('Absent' -eq $Ensure) -and ('Present' -eq $CurrentEventType.Ensure))
    {
        # If the Event Type exists and it shouldn't, simply remove it;
        Remove-ComplianceRetentionEventType -Identity $Name -confirm:$false
        Remove-ComplianceRetentionEventType -Identity $Name -confirm:$false -forcedeletion
    }
}

function Test-TargetResource
{
    [CmdletBinding()]
    [OutputType([System.Boolean])]
    param
    (
        [Parameter(Mandatory = $true)]
        [System.String]
        $Name,

        [Parameter()]
        [System.String]
        $Comment,

        [Parameter()]
        [System.DateTime]
        $EventDateTime,

        [Parameter()]
        [System.String[]]
        $EventTags,

        [Parameter()]
        [System.String[]]
        $EventTypes,

        [Parameter()]
        [Microsoft.Management.Infrastructure.CimInstance[]]
        $AssetId,

        [Parameter()]
        [Microsoft.Management.Infrastructure.CimInstance[]]
        $ExchangeAssetIdQuery,

        [Parameter()]
        [Microsoft.Management.Infrastructure.CimInstance[]]
        $SharePointAssetIdQuery,

        [Parameter()]
        [ValidateSet('Present', 'Absent')]
        [System.String]
        $Ensure = 'Present',

        [Parameter(Mandatory = $true)]
        [System.Management.Automation.PSCredential]
        $GlobalAdminAccount
    )
    #region Telemetry
    $ResourceName = $MyInvocation.MyCommand.ModuleName.Replace("MSFT_", "")
    $data = [System.Collections.Generic.Dictionary[[String], [String]]]::new()
    $data.Add("Resource", $ResourceName)
    $data.Add("Method", $MyInvocation.MyCommand)
    $data.Add("Principal", $GlobalAdminAccount.UserName)
    Add-M365DSCTelemetryEvent -Data $data
    #endregion

    Write-Verbose -Message "Testing configuration of RetentionComplianceEventType for $Name"

    $CurrentValues = Get-TargetResource @PSBoundParameters
    Write-Verbose -Message "Target Values: $(Convert-M365DscHashtableToString -Hashtable $PSBoundParameters)"

    $ValuesToCheck = $PSBoundParameters
    $ValuesToCheck.Remove('GlobalAdminAccount') | Out-Null

    $TestResult = Test-M365DSCParameterState -CurrentValues $CurrentValues `
        -Source $($MyInvocation.MyCommand.Source) `
        -DesiredValues $PSBoundParameters `
        -ValuesToCheck $ValuesToCheck.Keys

    Write-Verbose -Message "Test-TargetResource returned $TestResult"

    return $TestResult
}

function Export-TargetResource
{
    [CmdletBinding()]
    [OutputType([System.String])]
    param
    (
        [Parameter(Mandatory = $true)]
        [System.Management.Automation.PSCredential]
        $GlobalAdminAccount
    )

    #region Telemetry
    $ResourceName = $MyInvocation.MyCommand.ModuleName.Replace("MSFT_", "")
    $data = [System.Collections.Generic.Dictionary[[String], [String]]]::new()
    $data.Add("Resource", $ResourceName)
    $data.Add("Method", $MyInvocation.MyCommand)
    $data.Add("Principal", $GlobalAdminAccount.UserName)
    Add-M365DSCTelemetryEvent -Data $data
    #endregion

    $ConnectionMode = New-M365DSCConnection -Platform 'SecurityComplianceCenter' `
        -InboundParameters $PSBoundParameters `
        -SkipModuleReload $true

    try
    {
        $EventTypes = Get-ComplianceRetentionEventType -ErrorAction Stop
        $dscContent = ''

        Write-Host "`r`n" -NoNewline
        $i = 1

        foreach ($eventType in $EventTypes)
        {
            Write-Host "        |---[$i/$($EventTypes.Length)] $($eventType.Name)" -NoNewline

            $Params = @{
                GlobalAdminAccount = $GlobalAdminAccount
                Name               = $eventType.Name
            }
            $Results = Get-TargetResource @Params
            $Results = Update-M365DSCExportAuthenticationResults -ConnectionMode $ConnectionMode `
                -Results $Results
            $dscContent += Get-M365DSCExportContentForResource -ResourceName $ResourceName `
                -ConnectionMode $ConnectionMode `
                -ModulePath $PSScriptRoot `
                -Results $Results `
                -GlobalAdminAccount $GlobalAdminAccount
            Write-Host $Global:M365DSCEmojiGreenCheckMark
            $i++
        }
        return $dscContent
    }
    catch
    {
        try
        {
            Write-Verbose -Message $_
            $tenantIdValue = ""
            if (-not [System.String]::IsNullOrEmpty($TenantId))
            {
                $tenantIdValue = $TenantId
            }
            elseif ($null -ne $GlobalAdminAccount)
            {
                $tenantIdValue = $GlobalAdminAccount.UserName.Split('@')[1]
            }
            Add-M365DSCEvent -Message $_ -EntryType 'Error' `
                -EventID 1 -Source $($MyInvocation.MyCommand.Source) `
                -TenantId $tenantIdValue
        }
        catch
        {
            Write-Verbose -Message $_
        }
        return ""
    }
}

Export-ModuleMember -Function *-TargetResource
