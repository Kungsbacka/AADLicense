. "$PSScriptRoot\DynamicParameter.ps1"

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

function MakeAssignedLicense($InputObject)
{
    $license = New-Object 'System.Collections.Generic.List[Microsoft.Open.AzureAD.Model.AssignedLicense]'
    foreach ($item in $InputObject)
    {
        $license.Add((New-Object -TypeName 'Microsoft.Open.AzureAD.Model.AssignedLicense' -ArgumentList @($item.DisabledPlans, $item.SkuId)))
    }
    $license
}


function ConvertTo-AADAssignedLicense
{
<#
.SYNOPSIS
    Converts an object or JSON string into an Microsoft.Open.AzureAD.Model.AssignedLicenses object.

.DESCRIPTION
    The function takes two kinds of custom objects as input and converts them to an AssignedLicenses
    object that can be used with the AzureAD module.

    If the input object is a string it is first deserialized as JSON.

    The first type of object should contain only two properties: SkuId and DisabledPlans. SkuId is
    the license SkuId GUID and DisabledPlans is an array with ServicePlanId GUIDs that should be
    disabled. This will result in an AssignedLicenses object with only the AddLicenses property set.

    The second object should contain an AddLicenses property and a RemoveLicenses property. Both
    properties can contain one or more licenses of the first type. DisabledPlans property is ignored
    for RemoveLicenses.

    All functions in this module uses a custom object to describe and pass around license information.
    You only have to use this function if you want to use the AzureAD module to add or remove licenses.

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

.EXAMPLE

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
        $output = New-Object -TypeName 'Microsoft.Open.AzureAD.Model.AssignedLicenses'
        if ($object.AddLicenses -or $object.RemoveLicenses)
        {
            $output.AddLicenses = MakeAssignedLicense $object.AddLicenses
            $output.RemoveLicenses = [string[]]($object.RemoveLicenses.SkuId)   
        }
        else
        {
            $output.AddLicenses = MakeAssignedLicense $object
        }
        Write-Output -InputObject $output
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
        [string]
        $ObjectId,
        [Parameter(Mandatory=$true,ParameterSetName='All')]
        [switch]
        $All
    )
    DynamicParam
    {
        if (-not $All)
        {
            if ($ObjectId)
            {
                $objId = $ObjectId.Trim("'", '"')
                try
                {
                    $aadUser = Get-AzureADUser -ObjectId $objId
                }
                catch
                {
                }
            }
            $set = [System.Collections.Generic.SortedSet[string]]::new()
            foreach ($lic in $Script:LicenseBySkuId.Keys)
            {
                if ($aadUser -and $lic -notin $aadUser.AssignedLicenses.SkuId)
                {
                    continue
                }
                [void]$set.Add($Script:LicenseBySkuId[$lic])
            }
            if ($set.Count -eq 0)
            {
                [void]$set.Add('(User has no license)')
            }
            ([DynamicParameter]@{
                Name = 'License'
                Type = [string[]]
                Mandatory = $true
                ParameterSetName = 'Selected'
                ValidateSet = $set
            }).Get()
        }
    }
    process
    {
        $License = $PSBoundParameters['License']
        $licensesToRemove = @{}
        if ($License)
        {
            if ($License -eq '(User has no license)')
            {
                Write-Warning "User with ID ""$ObjectId"" is unlicensed."
                return
            }
            foreach ($name in $License)
            {
                $licensesToRemove.Add($Script:LicenseByName[$name], $true)
            }
        }
        $aadUser = Get-AzureADUser -ObjectId $ObjectId
        $removeLicenses = New-Object 'System.Collections.Generic.List[string]'
        foreach ($item in $aadUser.AssignedLicenses)
        {
            if ($All -or $licensesToRemove[$item.SkuId])
            {
                $removeLicenses.Add($item.SkuId)
            }
        }
        if ($removeLicenses.Count -eq 0)
        {
            Write-Warning -Message "User with ID ""$ObjectId"" is not assigned any of the licenses requested for removal."
            return
        }
        if ($PSCmdlet.ShouldProcess($ObjectId, 'Remove license(s)'))
        {
            $params = @{
                ObjectId = $ObjectId
                AssignedLicenses = New-Object -TypeName 'Microsoft.Open.AzureAD.Model.AssignedLicenses' -ArgumentList @($null, $removeLicenses)
            }
            Set-AzureADUserLicense @params
        }
    }
}

function Show-AADLicense
{
    param
    (
        [Parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true)]
        [object[]]
        $License
    )
    begin
    {
        $disabledPlans = [System.Collections.Generic.SortedSet[string]]::new()
    }
    process
    {
        foreach ($item in $License)
        {
            $displayName = $Script:LicenseBySkuId[$item.SkuId]
            if (-not $displayName)
            {
                Write-Warning "Unknown license $($item.SkuId)"
                $displayName = $item.SkuId
            }
            $disabledPlans.Clear()
            foreach ($plan in $item.DisabledPlans)
            {
                $planDisplayName = $Script:ServicePlanByServicePlanId[$plan]
                if (-not $planDisplayName)
                {
                    Write-Warning "Unknown service plan $plan"
                    $planDisplayName = $plan
                }
                [void]$disabledPlans.Add($planDisplayName)
            }
            Write-Host -Object "`r`n$displayName" -ForegroundColor Yellow
            if ($disabledPlans.Count -gt 0)
            {
                Write-Host -Object ('  [Disabled] ' + ($disabledPlans -join "`r`n  [Disabled] "))
            }
            else
            {
                Write-Host -Object '  [No disabled plans]'
            }
            Write-Host
        }
    }
}

$getAADLicenseFunc = @'
function Script:New-AADLicense
{
<#
    .SYNOPSIS
    Creates a new license object
    
    .DESCRIPTION
    Creates a new license object that can be used with the other function in this module
    that takes a license.

    Use ConvertTo-AADAssignedLicense to convert the license object into an AssignedLicenses
    object that can be used directly with the AzureAD module.
    
    License information is stored in license.json in the module folder. Information about
    service plans is stored in serviceplan.json.
    
    Underscores are added to all license names to get tab completion work with the dynamic
    parameter. If a static parameter has a validate set with values that contain spaces,
    and the static parameter is specified first, the dynamic parameter won't show up when
    you try to do tab completion.

    .PARAMETER Name
    License name
    
    .PARAMETER DisabledPlan
    Disabled service plans
    
    .EXAMPLE
    Get-AADLicense -Name Office_365_Enterprise_E3 -DisabledPlan 'Exchange Online (Plan 2)'

    .EXAMPLE
    If you need to set more than one license, you add them to an array:
    $licenses = @()
    $licenses += Get-AADLicense -Name Office_365_Enterprise_E3 -DisabledPlan 'Exchange Online (Plan 2)'
    $licenses += Get-AADLicense -Name Power_BI_Pro
    
    .NOTES
    License information in license.json and serviceplan.json is by no means a complete list of
    all licenses and service plans in Office 365/Azure. It contains everything I have found in our
    own Office 365 tennant. License and service plan names change all the time, but I will try to
    keep the list up to date as long as I use this module.
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

# Dynamically create function Get-AADLicense. This is done so that the dynamic parameter
# DisabledPlan can depend on the value of the License parameter.
$set = $Script:LicenseByName.Keys -join '","' -replace ' ', '_'
$func = $getAADLicenseFunc -replace '<SET>', $set
[scriptblock]::Create($func).Invoke()

# Export only the functions using PowerShell standard verb-noun naming.
# Be sure to list each exported functions in the FunctionsToExport field of the module manifest file.
# This improves performance of command discovery in PowerShell.
Export-ModuleMember -Function *-*
