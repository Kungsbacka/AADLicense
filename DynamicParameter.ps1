using namespace System.Management.Automation

class DynamicParameter
{
    [string]$Name
    [type]$Type = [string]
    [int]$Position = 0x80000000
    [bool]$Mandatory = $false
    [bool]$ValueFromPipeline = $false
    [bool]$ValueFromPipelineByPropertyName = $false
    [bool]$ValueFromRemainingArguments = $false
    [string]$ParameterSetName = '__AllParameterSets'
    [string]$HelpMessage = $null
    [string]$HelpMessageBaseName = $null
    [string]$HelpMessageResourceId = $null
    [string[]]$ValidateSet = $null

    [object] Get()
    {
        $paramDict = [RuntimeDefinedParameterDictionary]::new()
        return $this.Get($paramDict)
    }

    [object] Get($paramDict)
    {
        if ([string]::IsNullOrEmpty($this.Name))
        {
            throw 'Name cannot be null och empty'
        }
        $attribCol = [System.Collections.ObjectModel.Collection[System.Attribute]]::new()
        $paramAttrib = [ParameterAttribute]::new()
        $paramAttrib.Mandatory = $this.Mandatory
        $paramAttrib.Position = $this.Position
        $paramAttrib.ParameterSetName = $this.ParameterSetName
        $paramAttrib.ValueFromPipeline = $this.ValueFromPipeline
        $paramAttrib.ValueFromPipelineByPropertyName = $this.ValueFromPipelineByPropertyName
        $paramAttrib.ValueFromRemainingArguments = $this.ValueFromRemainingArguments
        $nullOrEmptyNotAllowed = @('HelpMessage','HelpMessageBaseName','HelpMessageResourceId')
        foreach ($item in $nullOrEmptyNotAllowed)
        {
            if (-not [string]::IsNullOrEmpty($this.$item))
            {
                $paramAttrib.$item = $this.$item
            }
        }
        if ($this.ValidateSet)
        {
            $validateSetAttrib = [System.Management.Automation.ValidateSetAttribute]::new($this.ValidateSet)
            $attribCol.Add($validateSetAttrib)
        }
        $attribCol.Add($paramAttrib)
        $param = [RuntimeDefinedParameter]::new($this.Name, $this.Type, $attribCol)
        $paramDict.Add($this.Name, $param)
        return $paramDict
    }
}
