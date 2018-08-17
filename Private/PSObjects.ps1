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

function Get-ObjectTypeInside {
    [CmdletBinding()]
    param(
        $Object
    )
    if ($Object -ne $null) {
        $ObjectType = $Object.GetType().Name
        if ($ObjectType -eq 'Object[]') {
            if ((Get-ObjectCount $Object) -gt 0) {
                $ObjectTypeInsider = $Object[0].GetType().Name

            }
        } else {
            $ObjectTypeInsider = $ObjectType
        }
    }

    return $ObjectTypeInsider
}
function Get-ObjectType {
    [CmdletBinding()]
    param(
        [Object] $Object,
        [string] $ObjectName = 'Random Object Name'
    )
    $Return = [ordered] @{}
    $Return.ObjectName = $ObjectName

    if ($Object -ne $null) {
        $TypeInformation = $Object.GetType()

        $Return.ObjectTypeName = $TypeInformation.Name
        $Return.ObjectTypeBaseName = $TypeInformation.BaseType
        $Return.SystemType = $TypeInformation.UnderlyingSystemType

        if ((Get-ObjectCount $Object) -gt 0) {
            #Write-Verbose "Get-ObjectType - $($Object.Count)"
            $TypeInformationInsider = $Object[0].GetType()
            $Return.ObjectTypeInsiderName = $TypeInformationInsider.Name
            $Return.ObjectTypeInsiderBaseName = $TypeInformationInsider.BaseType
            $Return.SystemTypeInsider = $TypeInformationInsider.UnderlyingSystemType
        } else {
            $Return.ObjectTypeInsiderName = ''
            $Return.ObjectTypeInsiderBaseName = ''
            $Return.SystemTypeInsider = ''
        }
    } else {
        $Return.ObjectTypeName = ''
        $Return.ObjectTypeBaseName = ''
        $Return.ObjectTypeInsiderName = ''
        $Return.ObjectTypeInsiderBaseName = ''
        $Return.SystemTypeInsider = ''

    }
    return  $Return.ForEach( {[PSCustomObject]$_})
}