function Format-PSTableConvertType3 {
    [CmdletBinding()]
    param (
        $Object
    )
    Write-Verbose 'Format-PSTableConvertType3 - Option 3'
    $Array = New-ArrayList
    ### Add Titles
    $Titles = New-ArrayList
    Add-ToArray -List $Titles -Element 'Name'
    Add-ToArray -List $Titles -Element 'Value'
    Add-ToArray -List $Array -Element $Titles

    ### Add Data
    foreach ($O in $Object) {
        foreach ($Key in $O.Keys) {
            # Write-Verbose "Test2 - $Key - $($O[$Key])"
            $ArrayValues = New-ArrayList
            Add-ToArray -List $ArrayValues -Element $Key
            Add-ToArray -List $ArrayValues -Element $O[$Key]
            Add-ToArray -List $Array -Element $ArrayValues
        }
    }
    return , $Array
}
function Format-PSTableConvertType2 {
    [CmdletBinding()]
    param(
        $Object
    )
    Write-Verbose 'Format-PSTableConvertType2 - Option 2'
    $Array = New-ArrayList
    ### Add Titles
    $Titles = New-ArrayList
    foreach ($O in $Object) {
        foreach ($Name in $O.PSObject.Properties.Name) {
            #Write-Verbose $Name
            Add-ToArray -List $Titles -Element $Name
        }
        break
    }
    Add-ToArray -List ($Array) -Element $Titles
    ### Add Data
    foreach ($O in $Object) {
        $ArrayValues = New-ArrayList
        foreach ($Value in $O.PSObject.Properties.Value) {
            #Write-Verbose $Value
            Add-ToArray -List $ArrayValues -Element $Value
        }
        Add-ToArray -List $Array -Element $ArrayValues
    }
    return , $Array
}
function Format-PSTableConvertType1 {
    [CmdletBinding()]
    param ($Object)
    Write-Verbose 'Format-PSTableConvertType1 - Option 1'
    $Array = New-ArrayList
    ### Add Titles
    # $Array += , @('Name', 'Value')
    $Titles = New-ArrayList
    Add-ToArray -List $Titles -Element 'Name'
    Add-ToArray -List $Titles -Element 'Value'
    Add-ToArray -List $Array -Element $Titles
    ### Add Data
    foreach ($Key in $Object.Keys) {
        Write-Verbose $Key
        Write-Verbose $Object.$Key
        #$Array += , @($Key, $Object.$Key)
        $ArrayValues = New-ArrayList
        Add-ToArray -List $ArrayValues -Element $Key
        Add-ToArray -List $ArrayValues -Element $Object.$Key
        Add-ToArray -List $Array -Element $ArrayValues
    }

    return , $Array
}


function Format-PSTable {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)] $Object
    )

    $Type = Get-ObjectType -Object $Object
    Write-Verbose "Format-PSTable - Type: $($Type.ObjectTypeName)"

    if ($Type.ObjectTypeName -eq 'Object[]' -or $Type.ObjectTypeName -eq 'OrderedDictionary' -or
        $Type.ObjectTypeName -eq 'Object' -or $Type.ObjectTypeName -eq 'PSCustomObject' -or
        $Type.ObjectTypeName -eq 'Collection`1') {

        if ($Type.ObjectTypeInsiderName -eq 'string') {
            return Format-PSTableConvertType1 -Object $Object
        } elseif ($Type.ObjectTypeInsiderName -eq 'Object' -or $Type.ObjectTypeInsiderName -eq 'PSCustomObject') {
            return Format-PSTableConvertType2 -Object $Object
        } elseif ($Type.ObjectTypeInsiderName -eq 'HashTable' -or $Type.ObjectTypeInsiderName -eq 'OrderedDictionary' ) {
            return Format-PSTableConvertType3 -Object $Object
        } else {
            # Covers ADDriveInfo and other types of objects
            return Format-PSTableConvertType2 -Object $Object
        }
    } elseif ($Type.ObjectTypeName -eq 'HashTable') {
        return Format-PSTableConvertType3 -Object $Object
    } else {
        throw 'Not supported? Weird'
    }
}