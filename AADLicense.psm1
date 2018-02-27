﻿. "$PSScriptRoot\DynamicParameter.ps1"

function Script:CheckResources()
{
    $notFound = 'license.json', 'serviceplan.json', 'licensepack.json' |
        Where-Object {-not (Test-Path -Path "$PSScriptRoot\$($_)")}
    if ($notFound)
    {
        Write-Error -Message 'Unable to find one or more resource files (check TargetObject).' -TargetObject $notFound
    }
}

function Script:LoadResources()
{
    # Load Licenses
    $licenses = Get-Content -Path "$PSScriptRoot\license.json" | ConvertFrom-Json
    $Script:LicenseBySkuId = @{}
    $Script:LicenseByName = @{}
    foreach ($item in $licenses)
    {
        $Script:LicenseBySkuId.Add($item.SkuId, ($item.LicenseName -replace ' ', '_'))
        $Script:LicenseByName.Add(($item.LicenseName -replace ' ', '_'), $item.SkuId)
    }

    # Load service plans
    $servicePlans = Get-Content -Path "$PSScriptRoot\serviceplan.json" | ConvertFrom-Json
    $Script:ServicePlanBySkuId = @{}
    $Script:ServicePlanByName = @{}
    $Script:ServicePlanByServicePlanId = @{}
    foreach ($item in $servicePlans)
    {
        if ($Script:ServicePlanBySkuId.ContainsKey($item.SkuId))
        {
            $Script:ServicePlanBySkuId[$item.SkuId] += $item
        }
        else
        {
           $Script:ServicePlanBySkuId.Add($item.SkuId, @($item))
        }
        if (-not $Script:ServicePlanByName.ContainsKey($item.Name))
        {
            $Script:ServicePlanByName.Add($item.Name, $item.ServicePlanId)
        }
        if (-not $Script:ServicePlanByServicePlanId.ContainsKey($item.ServicePlanId))
        {
            $Script:ServicePlanByServicePlanId.Add($item.ServicePlanId, $item.Name)
        }
    }
    
    # Load license packs
    $Script:LicensePacks = Get-Content -Path "$PSScriptRoot\licensepack.json" | ConvertFrom-Json
}

function ConvertTo-AADAssignedLicense
{
<#
.SYNOPSIS
    Converts an object or JSON string into an Microsoft.Open.AzureAD.Model.AssignedLicenses object.

.DESCRIPTION
    The function takes an input object with two properties: SkuId and DisabledPlans. SkuId is
    the license SkuId GUID and DisabledPlans is an array with ServicePlanId GUIDs that should
    be disabled. If the input object is a string it is first deserialized as JSON.

    All functions in this module uses this custom object format to describe and pass around license
    information. If you want to work with the AzureAD cmdlets directly you can use this function
    to convert the custom object to the type used by the AzureAD module.

.PARAMETER InputObject
Object or string to convert

.EXAMPLE
$obj = [pscustomobject]@{
    SkuId = '34ca1328-4568-40df-a09e-a5ab5e2c30e2'
    DisabledPlans = @('fc01bf39-abbb-44b4-8fbb-e69b1c0dff66','8c53903a-e937-40b5-b056-29cc51d902cd')
}
$obj | ConvertTo-AADAssignedLicense

.EXAMPLE
$json = '{"SkuId":"34ca1328-4568-40df-a09e-a5ab5e2c30e2","DisabledPlans":["fc01bf39-abbb-44b4-8fbb-e69b1c0dff66","8c53903a-e937-40b5-b056-29cc51d902cd"]}'
}
$obj | ConvertTo-AADAssignedLicense

.NOTES
Only the AddLicenses member is populated. RemoveLicenses is always empty.
#>
    param
    (
        [Parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true)]
        $InputObject
    )
    process
    {
        if ($InputObject -is [string])
        {
            $object = ConvertFrom-Json -InputObject $InputObject
        }
        else
        {
            $object = $InputObject
        }
        $licenses = New-Object 'System.Collections.Generic.List[Microsoft.Open.AzureAD.Model.AssignedLicense]'
        foreach ($item in $object)
        {
            $licenses.Add((New-Object -TypeName 'Microsoft.Open.AzureAD.Model.AssignedLicense' -ArgumentList @($item.DisabledPlans, $item.SkuId)))
        }
        New-Object -TypeName 'Microsoft.Open.AzureAD.Model.AssignedLicenses' -ArgumentList @($licenses, $null)
    }
}

function Get-AADLicensePack
{
<#
.SYNOPSIS
Gets a license pack by id or name.

.DESCRIPTION
License packs are a way to package one or more liceses with optional disabled service plans.
Packages are currently stored in a file called licensepack.json found in the module folder.
The license packs are given an ID and a name that can be used with this function.

.PARAMETER Id
License pack ID

.PARAMETER Name
License pack name

.EXAMPLE
Get-AADLicensePack -Name 'Global admins'

.NOTES
ConvertTo-AADAssignedLicense can be used to convert a license pack to an AssignedLicenses object.
#>
    [CmdletBinding()]
    param
    (
    )
    DynamicParam
    {
        $params = ([DynamicParameter]@{
            Name = 'Id'
            Mandatory = $true
            ParameterSetName = 'ById'
            ValidateSet = ($Script:LicensePacks.Id)
        }).Get()
        $params = ([DynamicParameter]@{
            Name = 'Name'
            Mandatory = $true
            ParameterSetName = 'ByName'
            ValidateSet = ($Script:LicensePacks.Name)
        }).Get($params)
        $params
    }
    begin
    {
        if ($PSBoundParameters.ContainsKey('Id'))
        {
            $pack = $Script:LicensePacks | Where-Object Id -EQ ($PSBoundParameters['Id'])
            Write-Output -InputObject $pack.AssignedLicenses
        }
        else
        {
            $pack = $Script:LicensePacks | Where-Object Name -EQ ($PSBoundParameters['Name'])
            Write-Output -InputObject $pack.AssignedLicenses
        }
    }
}

function Set-AADLicense
{
<#
.SYNOPSIS
Sets new licenses replacing all other licenses.

.DESCRIPTION
Sets new liceses for a user. The new licenses replaces all existing licenses.

If UsageLocation is not set on a user, assigning a license might fail. By supplying
the DefultUsageLocation parameter, the function can first set a UsageLocation before
licenses are assigned.

.PARAMETER ObjectId
ObjectId passed unchanged to AzureAD cmdlets

.PARAMETER License
Custom license object (see ConvertTo-AADAssignedLicense)

.PARAMETER DefaultUsageLocation
Default usage location string

.EXAMPLE
$license = Get-AADLicense -Name Office_365_Enterprise_E3 -DisabledPlan 'Microsoft StaffHub'
Get-AzureADUser -Filter "Department eq 'HR'" | Set-AADLicense -License $license -DefaultUsageLocation 'US'
#>
    [CmdletBinding(SupportsShouldProcess=$true)]
    param
    (
        [Parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        [string[]]
        $ObjectId,
        [Parameter(Mandatory=$true,Position=1,ValueFromPipelineByPropertyName=$true)]
        [object]
        $License,
        [string]
        $DefaultUsageLocation
    )
    process
    {
        foreach ($id in $ObjectId)
        {
            $aadUser = Get-AzureADUser -ObjectId $id
            if (-not $aadUser.UsageLocation -and $DefaultUsageLocation)
            {
                if ($PSCmdlet.ShouldProcess($aadUser.UserPrincipalName, 'Set UsageLocation'))
                {
                    Set-AzureADUser -ObjectId $id -UsageLocation $DefaultUsageLocation
                }
            }
            $removeLicenses = New-Object 'System.Collections.Generic.List[string]'
            foreach ($item in $aadUser.AssignedLicenses)
            {
                if ($item.SkuId -notin $License.SkuId)
                {
                    $removeLicenses.Add($item.SkuId)
                }
            }
            $addLicenses = New-Object 'System.Collections.Generic.List[Microsoft.Open.AzureAD.Model.AssignedLicense]'
            foreach ($item in $License)
            {
                $addLicenses.Add((New-Object -TypeName 'Microsoft.Open.AzureAD.Model.AssignedLicense' -ArgumentList @($item.DisabledPlans, $item.SkuId)))
            }
            $assignedLicenses = New-Object -TypeName 'Microsoft.Open.AzureAD.Model.AssignedLicenses' -ArgumentList @($addLicenses, $removeLicenses)
            if ($PSCmdlet.ShouldProcess($aadUser.UserPrincipalName, 'Replace license(s)'))
            {
                Set-AzureADUserLicense -ObjectId $id -AssignedLicenses $assignedLicenses
            }
        }
    }
}

function Add-AADLicense
{
<#
.SYNOPSIS
Adds a new license.

.DESCRIPTION
Adds new licenses to a user without touching existing licenses.

.PARAMETER ObjectId
ObjectId passed unchanged to AzureAD cmdlets

.PARAMETER License
Custom license object (see ConvertTo-AADAssignedLicense)

.PARAMETER DefaultUsageLocation
Default usage location string (see Set-AADLicense)

.EXAMPLE
see Set-AADLicense

.NOTES
This function is implemented for completness since it is almost as easy to use
Set-AzureADUserLicense directly in combination with Get-AADLicense/Get-AADLicensePack
and ConvertTo-AADAssignedLicense.
#>
    [CmdletBinding(SupportsShouldProcess=$true)]
    param
    (
        [Parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true)]
        [string[]]
        $ObjectId,
        [Parameter(Mandatory=$true,Position=1)]
        [object]
        $License,
        [string]
        $DefaultUsageLocation
    )
    begin
    {
        $newLicenses = $License | ConvertTo-AADAssignedLicense
    }
    process
    {
        foreach ($id in $ObjectId)
        {
            $aadUser = Get-AzureADUser -ObjectId $id
            if (-not $aadUser.UsageLocation -and $DefaultUsageLocation)
            {
                if ($PSCmdlet.ShouldProcess($aadUser.UserPrincipalName, 'Set UsageLocation'))
                {
                    Set-AzureADUser -ObjectId $id -UsageLocation $DefaultUsageLocation
                }
            }
            if ($PSCmdlet.ShouldProcess($id, 'Add license(s)'))
            {
                $params = @{
                    ObjectId = $id
                    AssignedLicenses = $newLicenses
                }
                Set-AzureADUserLicense @params
            }
        }
    }
}

function Remove-AADLicense
{
    [CmdletBinding(SupportsShouldProcess=$true)]
    param
    (
        [Parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true)]
        [string[]]
        $ObjectId,
        [Parameter(Mandatory=$true,ParameterSetName='AllLicenses')]
        [switch]
        $All
    )
    DynamicParam
    {
        # To be implemented
    }
    begin
    {
    }
    process
    {
        throw "Not implemented"
#        foreach ($id in $ObjectId)
#        {
#            $aadUser = Get-AzureADUser -ObjectId $id
#            if (-not $aadUser.AssignedLicenses)
#            {
#                Write-Warning -Message "User with ID ""$id"" has no licenses assigned. User is skipped."
#                continue
#            }
#            $removeLicenses = New-Object 'System.Collections.Generic.List[string]'
#            foreach ($item in $aadUser.AssignedLicenses)
#            {
#                if ($All -or $SkusToRemove.ContainsKey($item.SkuId))
#                {
#                    $removeLicenses.Add($item.SkuId)
#                }
#            }
#            if ($removeLicenses.Count -eq 0)
#            {
#                Write-Warning -Message "User with ID ""$id"" is not assigned any of the licenses requested for removal. User is skipped."
#                continue
#            }
#            if ($PSCmdlet.ShouldProcess($id, 'Remove license(s)'))
#            {
#                $params = @{
#                    ObjectId = $id
#                    AssignedLicenses = New-Object -TypeName 'Microsoft.Open.AzureAD.Model.AssignedLicenses' -ArgumentList @($null, $removeLicenses)
#                }
#                Set-AzureADUserLicense @params
#            }
#        }
    }
}

$getAADLicenseFunc = @'
function Script:Get-AADLicense
{
<#
    .SYNOPSIS
    Gets a license by name
    
    .DESCRIPTION
    This function makes working with licenses with the AzureAD module more convenient since
    you can reference licenses by their display name instead of by their GUID. You can also
    reference service plans by display name when disabling plans.
    
    License information is stored in license.json in the module folder. Information about
    service plans is stored in serviceplan.json.
    
    Underscores are added to all license names to get parameter binding to work as desired.
    If there are spaces in the names, parameter binding happens in a wat that the dynamic
    parameter (DisabledLicense) will not be able to read the license name. I have not found
    a workaround for this.

    .PARAMETER Name
    License name
    
    .PARAMETER Name
    License pack name
    
    .EXAMPLE
    Get-AADLicense -Name Office_365_Enterprise_E3 -DisabledPlan 'Exchange Online (Plan 2)'

    .EXAMPLE
    If you need to set more than one license, you add them to an array:
    $licenses = @()
    $licenses += Get-AADLicense -Name Office_365_Enterprise_E3 -DisabledPlan 'Exchange Online (Plan 2)'
    $licenses += Get-AADLicense -Name Power_BI_Pro
    
    .NOTES
    License information in license.json and serviceplan.json is by no means a complete list of
    all licenses and service plans in Office 365/Azure. It contains everything I could have found
    in our own Office 365 tennant. License and service plan names change all the time, but I will
    try to keep the list up to date as long as I use this module myself. I welcome suggestions of
    how to automate this process.
#>
        [CmdletBinding()]
    param
    (
        [ValidateSet("<SET>")]
        [Parameter(Mandatory=$true,Position=0)]
        [string]
        $Name
    )
    DynamicParam
    {
        $name = $PSBoundParameters['Name']
        if (-not $name)
        {
            return
        }
        $skuId = $Script:LicenseByName[$name]
        $servicePlans = $Script:ServicePlanBySkuId[$skuId]
        $set = $servicePlans | % Name
        ([DynamicParameter]@{
            Name = 'DisabledPlan'
            Type = [string[]]
            Position = 1
            ValidateSet = $set
        }).Get()
    }

    begin
    {
        $obj = [pscustomobject]@{
            SkuId = $Script:LicenseByName[$Name]
            DisabledPlans = $null
        }
        $disabledPlans = $PSBoundParameters['DisabledPlan']
        if ($disabledPlans)
        {
            $planIds = @()
            foreach ($item in $disabledPlans)
            {
                $planIds += $Script:ServicePlanByName[$item]
            }
            $obj.DisabledPlans = $planIds
        }
        Write-Output $obj
    }
}
'@


Script:CheckResources
Script:LoadResources

$set = $Script:LicenseByName.Keys -join '","' -replace ' ', '_'
$func = $getAADLicenseFunc -replace '<SET>', $set
[scriptblock]::Create($func).Invoke()

# Export only the functions using PowerShell standard verb-noun naming.
# Be sure to list each exported functions in the FunctionsToExport field of the module manifest file.
# This improves performance of command discovery in PowerShell.
Export-ModuleMember -Function *-*