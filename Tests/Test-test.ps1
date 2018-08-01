#Requires -Modules Pester
Import-Module $PSScriptRoot\..\PSWriteWord.psd1 #-Force
Clear-Host
### Preparing Data Start
$myitems0 = @(
    [pscustomobject]@{name = "Joe"; age = 32; info = "Cat lover"},
    [pscustomobject]@{name = "Sue"; age = 29; info = "Dog lover"},
    [pscustomobject]@{name = "Jason"; age = 42; info = "Food lover"}
)

$myitems1 = @(
    [pscustomobject]@{name = "Joe"; age = 32; info = "Cat lover"}
)
$myitems2 = [PSCustomObject]@{
    name = "Joe"; age = 32; info = "Cat lover"
}

$InvoiceEntry1 = @{}
$InvoiceEntry1.Description = 'IT Services 1'
$InvoiceEntry1.Amount = '$200'

$InvoiceEntry2 = @{}
$InvoiceEntry2.Description = 'IT Services 2'
$InvoiceEntry2.Amount = '$300'

$InvoiceEntry3 = @{}
$InvoiceEntry3.Description = 'IT Services 3'
$InvoiceEntry3.Amount = '$288'

$InvoiceEntry4 = @{}
$InvoiceEntry4.Description = 'IT Services 4'
$InvoiceEntry4.Amount = '$301'

$InvoiceEntry5 = @{}
$InvoiceEntry5.Description = 'IT Services 5'
$InvoiceEntry5.Amount = '$299'

$InvoiceData1 = @()
$InvoiceData1 += $InvoiceEntry1
$InvoiceData1 += $InvoiceEntry2
$InvoiceData1 += $InvoiceEntry3
$InvoiceData1 += $InvoiceEntry4
$InvoiceData1 += $InvoiceEntry5

$InvoiceData2 = $InvoiceData1.ForEach( {[PSCustomObject]$_})

$InvoiceData3 = @()
$InvoiceData3 += $InvoiceEntry1

$InvoiceData4 = $InvoiceData3.ForEach( {[PSCustomObject]$_})
### Preparing Data End

$Object1 = Get-Process | Select-Object ProcessName, Handle, StartTime -First 5
$Object2 = Get-PSDrive
$Object3 = Get-PSDrive | Select-Object * -First 2
$Object4 = Get-PSDrive | Select-Object * -First 1

$obj = New-Object System.Object
$obj | Add-Member -type NoteProperty -name Name -Value "Ryan_PC"
$obj | Add-Member -type NoteProperty -name Manufacturer -Value "Dell"
$obj | Add-Member -type NoteProperty -name ProcessorSpeed -Value "3 Ghz"
$obj | Add-Member -type NoteProperty -name Memory -Value "6 GB"

$myObject2 = New-Object System.Object
$myObject2 | Add-Member -type NoteProperty -name Name -Value "Doug_PC"
$myObject2 | Add-Member -type NoteProperty -name Manufacturer -Value "HP"
$myObject2 | Add-Member -type NoteProperty -name ProcessorSpeed -Value "2.6 Ghz"
$myObject2 | Add-Member -type NoteProperty -name Memory -Value "4 GB"

$myObject3 = New-Object System.Object
$myObject3 | Add-Member -type NoteProperty -name Name -Value "Julie_PC"
$myObject3 | Add-Member -type NoteProperty -name Manufacturer -Value "Compaq"
$myObject3 | Add-Member -type NoteProperty -name ProcessorSpeed -Value "2.0 Ghz"
$myObject3 | Add-Member -type NoteProperty -name Memory -Value "2.5 GB"

$myArray1 = @($obj, $myobject2, $myObject3)
$myArray2 = @($obj)


$InvoiceEntry7 = [ordered]@{}
$InvoiceEntry7.Description = 'IT Services 4'
$InvoiceEntry7.Amount = '$301'

$InvoiceEntry8 = [ordered]@{}
$InvoiceEntry8.Description = 'IT Services 5'
$InvoiceEntry8.Amount = '$299'

$InvoiceDataOrdered1 = @()
$InvoiceDataOrdered1 += $InvoiceEntry7

$InvoiceDataOrdered2 = @()
$InvoiceDataOrdered2 += $InvoiceEntry7
$InvoiceDataOrdered2 += $InvoiceEntry8


$Array = @()
$Array += Get-ObjectType -Object $myitems0  -ObjectName '$myitems0'
$Array += Get-ObjectType -Object $myitems1  -ObjectName '$myitems1'
$Array += Get-ObjectType -Object $myitems2 -ObjectName '$myitems2'
$Array += Get-ObjectType -Object $InvoiceEntry1 -ObjectName '$InvoiceEntry1'
$Array += Get-ObjectType -Object $InvoiceData1  -ObjectName '$InvoiceData1'
$Array += Get-ObjectType -Object $InvoiceData2  -ObjectName '$InvoiceData2'
$Array += Get-ObjectType -Object $InvoiceData3  -ObjectName '$InvoiceData3'
$Array += Get-ObjectType -Object $InvoiceData4  -ObjectName '$InvoiceData4'
$Array += Get-ObjectType -Object $Object1  -ObjectName '$Object1'
$Array += Get-ObjectType -Object $Object2  -ObjectName '$Object2'
$Array += Get-ObjectType -Object $Object3  -ObjectName '$Object3'
$Array += Get-ObjectType -Object $Object4  -ObjectName '$Object4'
$Array += Get-ObjectType -Object $obj -ObjectName '$obj'
$Array += Get-ObjectType -Object $myArray1 -ObjectName '$myArray1'
$Array += Get-ObjectType -Object $myArray2 -ObjectName '$myArray2'
$Array += Get-ObjectType -Object $InvoiceEntry7 -ObjectName '$InvoiceEntry7'
$Array += Get-ObjectType -Object $InvoiceDataOrdered1 -ObjectName '$InvoiceDataOrdered1'
$Array += Get-ObjectType -Object $InvoiceDataOrdered2 -ObjectName '$InvoiceDataOrdered2'
$Array | Format-Table -AutoSize


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
    if ($Type.ObjectTypeName -eq 'Object[]' -or
        $Type.ObjectTypeName -eq 'OrderedDictionary' -or
        $Type.ObjectTypeName -eq 'Object' -or
        $Type.ObjectTypeName -eq 'PSCustomObject' -or
        $Type.ObjectTypeName -eq 'Collection`1'
    ) {
        if ($Type.ObjectTypeInsiderName -eq 'string') {
            return Format-PSTableConvertType1 -Object $Object
        } elseif ($Type.ObjectTypeInsiderName -eq 'Object' -or
            $Type.ObjectTypeInsiderName -eq 'PSCustomObject' -or
            $Type.ObjectTypeInsiderName -eq 'ADDriveInfo'

        ) {
            return Format-PSTableConvertType2 -Object $Object
        } elseif ($Type.ObjectTypeInsiderName -eq 'HashTable' -or $Type.ObjectTypeInsiderName -eq 'OrderedDictionary' ) {
            return Format-PSTableConvertType3 -Object $Object
        }
    } elseif ($Type.ObjectTypeName -eq 'HashTable') {
        return Format-PSTableConvertType3 -Object $Object
    }
    Write-Verbose 'Option Exit'

}

function Show-TableVisualization {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)] $Object
    )
    Write-Color "[i] This is how table looks like in Format-Table" -Color Yellow
    $Object | Format-Table -AutoSize
    $Data = Format-PSTable $Object #-Verbose
    $RowNr = 0
    Write-Color "[i] Presenting table after conversion" -Color Yellow
    foreach ($Row in $Data) {
        $ColumnNr = 0
        foreach ($Column in $Row) {
            Write-Color 'Row: ', $RowNr, ' Column: ', $ColumnNr, " Data: ", $Column -Color White, Yellow, White, Green
            $ColumnNr++
        }
        $RowNr++
    }
}

#Show-TableVisualization $myItems0 -Verbose
#Show-TableVisualization $myItems1 -Verbose
#Show-TableVisualization $myItems2 -Verbose
#Show-TableVisualization $InvoiceEntry1 -Verbose
#Show-TableVisualization $InvoiceData1 -Verbose
#Show-TableVisualization $InvoiceData2 -Verbose
#Show-TableVisualization $InvoiceData3 -Verbose
#Show-TableVisualization $InvoiceData4 -Verbose
#Show-TableVisualization $Object1 -Verbose
#Show-TableVisualization $Object2 -Verbose ### Seems to be really weird - $Object2 | fl *
#Show-TableVisualization $Object3 -Verbose ### Seems to be really weird - $Object3 | fl *
#Show-TableVisualization $Object4 -Verbose
#Show-TableVisualization $obj -Verbose
#Show-TableVisualization $myArray1 -Verbose
#Show-TableVisualization $myArray2 -Verbose
#Show-TableVisualization $InvoiceEntry7 -Verbose
#Show-TableVisualization $InvoiceDataOrdered1 -Verbose
#Show-TableVisualization $InvoiceDataOrdered2 -Verbose