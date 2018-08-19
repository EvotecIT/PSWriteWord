function Get-ObjectTitles {
    [CmdletBinding()]
    param(
        $Object
    )
    $ArrayList = New-Object System.Collections.ArrayList
    Write-Verbose "Get-ObjectTitles - ObjectType $($Object.GetType())"
    foreach ($Title in $Object.PSObject.Properties) {
        Write-Verbose "Get-ObjectTitles - Value added to array: $($Title.Name)"
        $ArrayList.Add($Title.Name) | Out-Null
    }
    Write-Verbose "Get-ObjectTitles - Array size: $($ArrayList.Count)"
    return $ArrayList
}
function Get-ObjectData {
    [CmdletBinding()]
    param(
        $Object,
        $Title,
        [switch] $DoNotAddTitles
    )
    $ArrayList = New-Object System.Collections.ArrayList
    $Values = $Object.$Title
    Write-Verbose "Get-ObjectData1: Title $Title Values: $Values"
    if ((Get-ObjectCount $values) -eq 1 -and $DoNotAddTitles -eq $false) {
        $ArrayList.Add("$Title - $Values") | Out-Null
    } else {
        if ($DoNotAddTitles -eq $false) { $ArrayList.Add($Title) | Out-Null }
        foreach ($Value in $Values) {
            $ArrayList.Add("$Value") | Out-Null
        }
    }
    Write-Verbose "Get-ObjectData2: Title $Title Values: $(Get-ObjectCount $ArrayList)"
    return $ArrayList
}
function Get-ObjectCount {
    [CmdletBinding()]
    param(
        $Object
    )
    return $($Object | Measure-Object).Count
}
function Get-ObjectType {
    [CmdletBinding()]
    param(
        [Object] $Object,
        [string] $ObjectName = 'Random Object Name'
    )
    $Data = [ordered] @{}
    $Data.ObjectName = $ObjectName

    if ($Object) {
        try {
            $TypeInformation = $Object.GetType()
            $Data.ObjectTypeName = $TypeInformation.Name
            $Data.ObjectTypeBaseName = $TypeInformation.BaseType
            $Data.SystemType = $TypeInformation.UnderlyingSystemType
        } catch {
            $Data.ObjectTypeName = ''
            $Data.ObjectTypeBaseName = ''
            $Data.SystemType = ''
        }
        try {
            $TypeInformationInsider = $Object[0].GetType()
            $Data.ObjectTypeInsiderName = $TypeInformationInsider.Name
            $Data.ObjectTypeInsiderBaseName = $TypeInformationInsider.BaseType
            $Data.SystemTypeInsider = $TypeInformationInsider.UnderlyingSystemType
        } catch {
            $Data.ObjectTypeInsiderName = ''
            $Data.ObjectTypeInsiderBaseName = ''
            $Data.SystemTypeInsider = ''
        }
    } else {
        $Data.ObjectTypeName = ''
        $Data.ObjectTypeBaseName = ''
        $Data.SystemType = ''
        $Data.ObjectTypeInsiderName = ''
        $Data.ObjectTypeInsiderBaseName = ''
        $Data.SystemTypeInsider = ''
    }
    return Format-TransposeTable -Object $Data
}