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

$obj | ft -a
#$obj.GetType() | fl *

$myObject2 = New-Object System.Object
$myObject2 | Add-Member -type NoteProperty -name Name -Value "Doug_PC"
$myObject2 | Add-Member -type NoteProperty -name Manufacturer -Value "HP"
$myObject2 | Add-Member -type NoteProperty -name ProcessorSpeed -Value "2.6 Ghz"
$myObject2 | Add-Member -type NoteProperty -name Memory -Value "4 GB"


$myObject2 | ft -a

$myObject3 = New-Object System.Object
$myObject3 | Add-Member -type NoteProperty -name Name -Value "Julie_PC"
$myObject3 | Add-Member -type NoteProperty -name Manufacturer -Value "Compaq"
$myObject3 | Add-Member -type NoteProperty -name ProcessorSpeed -Value "2.0 Ghz"
$myObject3 | Add-Member -type NoteProperty -name Memory -Value "2.5 GB"

$myArray1 = @($obj, $myobject2, $myObject3)
$myArray2 = @($obj)

$myArray1 | ft -a
$myArray2 | ft -a


$Array = @()

$Array += Get-ObjectType -Object $myitems0 #| ft -a
$Array += Get-ObjectType -Object $myitems1 #| ft -a
$Array += Get-ObjectType -Object $myitems2 #| ft -a
$Array += Get-ObjectType -Object $InvoiceEntry1 #| ft -a
$Array += Get-ObjectType -Object $InvoiceData1 #| ft -a
$Array += Get-ObjectType -Object $InvoiceData2 #| ft -a
$Array += Get-ObjectType -Object $InvoiceData3 #| ft -a
$Array += Get-ObjectType -Object $InvoiceData4 #| ft -a
# Not yet done

$Array += Get-ObjectType -Object $Object1 #| ft -a
$Array += Get-ObjectType -Object $Object2 #| ft -a
$Array += Get-ObjectType -Object $Object3 #| ft -a
$Array += Get-ObjectType -Object $Object4 #| ft -a

$Array += Get-ObjectType -Object $obj #| ft -a
$Array += Get-ObjectType -Object $myObject2 #| ft -a
$Array += Get-ObjectType -Object $myArray1 #| ft -a
$Array += Get-ObjectType -Object $myArray2 #| ft -a

#$Array | Sort-Object ObjectTypeName | Format-Table -AutoSize
$Array | Format-Table -AutoSize

Describe 'Add-WordTable - Should deliver same results as Format-Table -Autosize' {


    It 'Given Object[]/Array (MyItems0) with PSCustomObject should have 3 columns, 4 rows, 3rd row 3rd column should be Food lover' {
        <#  $myitems | Format-Table -AutoSize
        name  age info
        ----  --- ----
        Joe    32 Cat lover
        Sue    29 Dog lover
        Jason  42 Food lover
        #>
        $Type = Get-ObjectType -Object $myitems0
        $Type.ObjectTypeName | Should -Be 'Object[]'
        $Type.ObjectTypeBaseName | Should -Be 'Array'
        $Type.ObjectTypeInsiderName | Should -Be 'PSCustomObject'
        $Type.ObjectTypeInsiderBaseName | Should -Be 'System.Object'

        $WordDocument = New-WordDocument
        $WordDocument | Add-WordTable -DataTable $myitems0
        $WordDocument.Tables.Count | Should -Be 1
        $WordDocument.Tables[0].ColumnCount | Should -Be 3
        $WordDocument.Tables[0].RowCount | Should -Be 4
        $WordDocument.Tables[0].Rows[0].Cells[0].Paragraphs[0].Text | Should -Be 'name'
        $WordDocument.Tables[0].Rows[1].Cells[0].Paragraphs[0].Text | Should -Be 'Joe'
        $WordDocument.Tables[0].Rows[2].Cells[0].Paragraphs[0].Text | Should -Be 'Sue'
        $WordDocument.Tables[0].Rows[3].Cells[2].Paragraphs[0].Text | Should -Be 'Food lover'
    }
    It 'Given Object[]/Array (MyItems1) with PSCustomObject should have 3 columns, 2 rows, data should be in proper columns' {
        <#  $myitems1 | Format-Table -AutoSize
            name age info
            ---- --- ----
            Joe   32 Cat lover
        #>
        $Type = Get-ObjectType -Object $myitems1
        $Type.ObjectTypeName | Should -Be 'Object[]'
        $Type.ObjectTypeBaseName | Should -Be 'Array'
        $Type.ObjectTypeInsiderName | Should -Be 'PSCustomObject'
        $Type.ObjectTypeInsiderBaseName | Should -Be 'System.Object'

        $WordDocument = New-WordDocument
        $WordDocument | Add-WordTable -DataTable $myitems1
        $WordDocument.Tables.Count | Should -Be 1
        $WordDocument.Tables[0].ColumnCount | Should -Be 3
        $WordDocument.Tables[0].RowCount | Should -Be 2
        $WordDocument.Tables[0].Rows[0].Cells[0].Paragraphs[0].Text | Should -Be 'name'
        $WordDocument.Tables[0].Rows[1].Cells[0].Paragraphs[0].Text | Should -Be 'Joe'
        $WordDocument.Tables[0].Rows[1].Cells[1].Paragraphs[0].Text | Should -Be '32'
        $WordDocument.Tables[0].Rows[1].Cells[2].Paragraphs[0].Text | Should -Be 'Cat lover'
    }
    It 'Given PSCustomObject/System.Object (MyItems2) with PSCustomObject should have 3 columns, 2 rows, data should be in proper columns' {
        <#  $myitems2 | Format-Table -AutoSize
            name age info
            ---- --- ----
            Joe   32 Cat lover
        #>
        $Type = Get-ObjectType -Object $MyItems2
        $Type.ObjectTypeName | Should -Be 'PSCustomObject'
        $Type.ObjectTypeBaseName | Should -Be 'System.Object'
        $Type.ObjectTypeInsiderName | Should -Be 'PSCustomObject'
        $Type.ObjectTypeInsiderBaseName | Should -Be 'System.Object'

        $WordDocument = New-WordDocument
        $WordDocument | Add-WordTable -DataTable $myitems2
        $WordDocument.Tables.Count | Should -Be 1
        $WordDocument.Tables[0].ColumnCount | Should -Be 3
        $WordDocument.Tables[0].RowCount | Should -Be 2
        $WordDocument.Tables[0].Rows[0].Cells[0].Paragraphs[0].Text | Should -Be 'name'
        $WordDocument.Tables[0].Rows[1].Cells[0].Paragraphs[0].Text | Should -Be 'Joe'
        $WordDocument.Tables[0].Rows[1].Cells[1].Paragraphs[0].Text | Should -Be '32'
        $WordDocument.Tables[0].Rows[1].Cells[2].Paragraphs[0].Text | Should -Be 'Cat lover'
    }
    It 'Given Hashtable/System.Object (InvoiceEntry1) with HashTable should have 2 columns, 3 rows, data should be in proper columns' {
        <# $InvoiceEntry1 | Format-Table -AutoSize
            Name        Value
            ----        -----
            Description IT Services 1
            Amount      $200
            #>
        #
        $Type = Get-ObjectType -Object $InvoiceEntry1
        $Type.ObjectTypeName | Should -Be 'Hashtable'
        $Type.ObjectTypeBaseName | Should -Be 'System.Object'
        $Type.ObjectTypeInsiderName | Should -Be ''
        $Type.ObjectTypeInsiderBaseName | Should -Be ''

        $WordDocument = New-WordDocument
        $WordDocument | Add-WordTable -DataTable $InvoiceEntry1
        $WordDocument.Tables.Count | Should -Be 1
        $WordDocument.Tables[0].ColumnCount | Should -Be 2
        $WordDocument.Tables[0].RowCount | Should -Be 3
        $WordDocument.Tables[0].Rows[0].Cells[0].Paragraphs[0].Text | Should -Be 'Name'
        $WordDocument.Tables[0].Rows[1].Cells[0].Paragraphs[0].Text | Should -Be 'Description'
        $WordDocument.Tables[0].Rows[1].Cells[1].Paragraphs[0].Text | Should -Be 'IT Services 1'
        $WordDocument.Tables[0].Rows[2].Cells[1].Paragraphs[0].Text | Should -Be '$200'
    }
    It 'Given Object[]/Array (InvoiceData1) with HashTable should have 2 columns, 10 rows, data should be in proper columns' {
        <#
            Name        Value
            ----        -----
            Description IT Services 1
            Amount      $200
            Description IT Services 2
            Amount      $300
            Description IT Services 3
            Amount      $288
            Description IT Services 4
            Amount      $301
            Description IT Services 5
            Amount      $299

        #>
        #$InvoiceData1 | Format-Table -AutoSize
        $Type = Get-ObjectType -Object $InvoiceData1
        $Type.ObjectTypeName | Should -Be 'Object[]'
        $Type.ObjectTypeBaseName | Should -Be 'Array'
        $Type.ObjectTypeInsiderName | Should -Be 'Hashtable'
        $Type.ObjectTypeInsiderBaseName | Should -Be 'System.Object'

        $WordDocument = New-WordDocument
        $WordDocument | Add-WordTable -DataTable $InvoiceData1
        $WordDocument.Tables.Count | Should -Be 1
        $WordDocument.Tables[0].ColumnCount | Should -Be 2
        $WordDocument.Tables[0].RowCount | Should -Be 10
        $WordDocument.Tables[0].Rows[0].Cells[0].Paragraphs[0].Text | Should -Be 'Name'
        $WordDocument.Tables[0].Rows[1].Cells[0].Paragraphs[0].Text | Should -Be 'Description'
        $WordDocument.Tables[0].Rows[1].Cells[1].Paragraphs[0].Text | Should -Be 'IT Services 1'
        $WordDocument.Tables[0].Rows[2].Cells[1].Paragraphs[0].Text | Should -Be '$200'
    }
    It 'Given Collection`1/System.Object (InvoiceData2) with Collection`1 should have 2 columns, 6 rows, data should be in proper columns' {
        <#
            IsPublic IsSerial Name                                     BaseType
            -------- -------- ----                                     --------
            True     True     Collection`1                             System.Object
            Collection`1


            Description   Amount
            -----------   ------
            IT Services 1 $200
            IT Services 2 $300
            IT Services 3 $288
            IT Services 4 $301
            IT Services 5 $299

        #>
        #$InvoiceData1 | Format-Table -AutoSize
        $Type = Get-ObjectType -Object $InvoiceData2
        $Type.ObjectTypeName | Should -Be 'Collection`1'
        $Type.ObjectTypeBaseName | Should -Be 'System.Object'
        $Type.ObjectTypeInsiderName | Should -Be 'PSCustomObject'
        $Type.ObjectTypeInsiderBaseName | Should -Be 'System.Object'

        $WordDocument = New-WordDocument
        $WordDocument | Add-WordTable -DataTable $InvoiceData2
        $WordDocument.Tables.Count | Should -Be 1
        $WordDocument.Tables[0].ColumnCount | Should -Be 2
        $WordDocument.Tables[0].RowCount | Should -Be 6
        $WordDocument.Tables[0].Rows[0].Cells[0].Paragraphs[0].Text | Should -Be 'Description'
        $WordDocument.Tables[0].Rows[0].Cells[1].Paragraphs[0].Text | Should -Be 'Amount'
        $WordDocument.Tables[0].Rows[1].Cells[0].Paragraphs[0].Text | Should -Be 'IT Services 1'
        $WordDocument.Tables[0].Rows[2].Cells[1].Paragraphs[0].Text | Should -Be '$300'
    }
    It 'Given Object[]/Array (InvoiceData3) with HashTable should have 2 columns, 3 rows, data should be in proper columns' {
        <#
        IsPublic IsSerial Name                                     BaseType
        -------- -------- ----                                     --------
        True     True     Object[]                                 System.Array
        Hashtable



        Name        Value
        ----        -----
        Description IT Services 1
        Amount      $200

        #>
        #$InvoiceData1 | Format-Table -AutoSize
        $Type = Get-ObjectType -Object $InvoiceData3
        $Type.ObjectTypeName | Should -Be 'Object[]'
        $Type.ObjectTypeBaseName | Should -Be 'Array'
        $Type.ObjectTypeInsiderName | Should -Be 'Hashtable'
        $Type.ObjectTypeInsiderBaseName | Should -Be 'System.Object'

        $WordDocument = New-WordDocument
        $WordDocument | Add-WordTable -DataTable $InvoiceData3
        $WordDocument.Tables.Count | Should -Be 1
        $WordDocument.Tables[0].ColumnCount | Should -Be 2
        $WordDocument.Tables[0].RowCount | Should -Be 3
        $WordDocument.Tables[0].Rows[0].Cells[0].Paragraphs[0].Text | Should -Be 'Name'
        $WordDocument.Tables[0].Rows[0].Cells[1].Paragraphs[0].Text | Should -Be 'Value'
        $WordDocument.Tables[0].Rows[1].Cells[0].Paragraphs[0].Text | Should -Be 'Description'
        $WordDocument.Tables[0].Rows[1].Cells[1].Paragraphs[0].Text | Should -Be 'IT Services 1'
    }
    It 'Given Collection`1/System.Object (InvoiceData4) with Collection`1 should have 2 columns, 3 rows, data should be in proper columns' {
        <#
        IsPublic IsSerial Name                                     BaseType
        -------- -------- ----                                     --------
        True     True     Collection`1                             System.Object
        Collection`1



        Description   Amount
        -----------   ------
        IT Services 1 $200
        #>
        #$InvoiceData1 | Format-Table -AutoSize
        $Type = Get-ObjectType -Object $InvoiceData4
        $Type.ObjectTypeName | Should -Be 'Collection`1'
        $Type.ObjectTypeBaseName | Should -Be 'System.Object'
        $Type.ObjectTypeInsiderName | Should -Be 'PSCustomObject'
        $Type.ObjectTypeInsiderBaseName | Should -Be 'System.Object'

        $WordDocument = New-WordDocument
        $WordDocument | Add-WordTable -DataTable $InvoiceData4
        $WordDocument.Tables.Count | Should -Be 1
        $WordDocument.Tables[0].ColumnCount | Should -Be 2
        $WordDocument.Tables[0].RowCount | Should -Be 2
        $WordDocument.Tables[0].Rows[0].Cells[0].Paragraphs[0].Text | Should -Be 'Description'
        $WordDocument.Tables[0].Rows[0].Cells[1].Paragraphs[0].Text | Should -Be 'Amount'
        $WordDocument.Tables[0].Rows[1].Cells[0].Paragraphs[0].Text | Should -Be 'IT Services 1'
        $WordDocument.Tables[0].Rows[1].Cells[1].Paragraphs[0].Text | Should -Be '$200'
    }



}








Describe 'Add-WordTable - Should have proper settings' {
    It 'Given 2 tables, document should have 2 tables with proper design' {
        $WordDocument = New-WordDocument
        $Object2 = Get-PSDrive
        $WordDocument | Add-WordTable -DataTable $Object2 -Design 'ColorfulList' #-Verbose
        $WordDocument | Add-WordTable -DataTable $Object2 -Design "LightShading" #-Verbose
        $WordDocument.Tables[0].Design | Should -Be 'ColorfulList'
        $WordDocument.Tables[1].Design | Should -Be 'LightShading'
        $WordDocument.Tables.Count | Should -Be 2
    }
    It 'Given Array of PSCustomObject document should have 1 table with proper design, proper number of columns and rows' {
        $WordDocument = New-WordDocument

        Add-WordTable -WordDocument $WordDocument -DataTable $InvoiceData1 -Design MediumShading1 -AutoFit Contents #-Verbose
        $WordDocument.Tables[0].RowCount | Should -Be 6
        $WordDocument.Tables[0].ColumnCount | Should -Be 2
        # $WordDocument.Tables[0].AutoFit | Should -Be 'Contents' # Seems like a bug in Xceed - always returns ColumnWidth
        $WordDocument.Tables[0].Design | Should -Be 'MediumShading1'
    }
    It 'Given Array of PSCustomObejct document should have 1 table with proper design, proper number of columns and rows and proper index' {
        $WordDocument = New-WordDocument

        $InvoiceEntry1 = @{}
        $InvoiceEntry1.Description = 'IT Services 1'
        $InvoiceEntry1.Amount = '$200'

        $InvoiceData = @()
        $InvoiceData += $InvoiceEntry1

        Add-WordTable -WordDocument $WordDocument -DataTable $InvoiceData -Design ColorfulGrid
        $WordDocument.Tables[0].RowCount | Should -Be 2
        $WordDocument.Tables[0].ColumnCount | Should -Be 2
        $WordDocument.Tables[0].Index | Should -Be 0
        $WordDocument.Tables[0].Design | Should -Be 'ColorfulGrid'
    }
    It 'Given Array of 2 tables document should have 2 tables with proper row count, column count and design' {
        $WordDocument = New-WordDocument $FilePath
        $Object1 = Get-Process | Select-Object ProcessName, Handle, StartTime -First 5
        Add-WordTable -WordDocument $WordDocument -DataTable $Object1 -Design 'ColorfulList' -Supress $true #-Verbose
        $Object2 = Get-PSDrive | Select-Object * -First 2
        Add-WordTable -WordDocument $WordDocument -DataTable $Object2 -Design "LightShading" -MaximumColumns 7 -Supress $true #-Verbose

        $WordDocument.Tables[0].RowCount | Should -Be 6
        $WordDocument.Tables[0].ColumnCount | Should -Be 3
        $WordDocument.Tables[0].Design | Should -Be 'ColorfulList'
        $WordDocument.Tables[1].RowCount | Should -Be 3
        $WordDocument.Tables[1].ColumnCount | Should -Be 7
        $WordDocument.Tables[1].Design | Should -Be 'LightShading'
        $WordDocument.Tables.Count | Should -Be 2
    }
}
