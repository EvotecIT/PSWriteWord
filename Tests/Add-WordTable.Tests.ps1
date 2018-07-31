#Requires -Modules Pester
Import-Module $PSScriptRoot\..\PSWriteWord.psd1 #-Force


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

$InvoiceData2 = @()
$InvoiceData2 += $InvoiceEntry1

$InvoiceData3 = @()
$InvoiceData3 += $InvoiceEntry1
$InvoiceData3 += $InvoiceEntry2
### Preparing Data End


Clear-Host

$InvoiceEntry1.GetType()
Get-ObjectTypeInside -Object $InvoiceEntry1
$InvoiceEntry1 | ft -a

Describe 'Add-WordTable - Should deliver same results as Format-Table -Autosize' {

    It 'Given Object[]/Array (MyItems0) with PSCustomObject should have 3 columns, 4 rows, 3rd row 3rd column should be Food lover' {
        <#  $myitems | Format-Table -AutoSize
        name  age info
        ----  --- ----
        Joe    32 Cat lover
        Sue    29 Dog lover
        Jason  42 Food lover
        #>
        $myitems0.GetType().Name | Should -Be 'Object[]'
        $myitems0.GetType().BaseType | Should -Be 'Array'
        Get-ObjectTypeInside -Object $myitems0 | Should -Be 'PSCustomObject'

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
        $myitems1.GetType().Name | Should -Be 'Object[]'
        $myitems1.GetType().BaseType | Should -Be 'Array'
        Get-ObjectTypeInside -Object $myitems1 | Should -Be 'PSCustomObject'

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
        $myitems2.GetType().Name | Should -Be 'PSCustomObject'
        $myitems2.GetType().BaseType | Should -Be 'System.Object'
        Get-ObjectTypeInside -Object $myitems2 | Should -Be 'PSCustomObject'

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
        $InvoiceEntry1.GetType().Name | Should -Be 'Hashtable'
        $InvoiceEntry1.GetType().BaseType | Should -Be 'System.Object'
        Get-ObjectTypeInside -Object $InvoiceEntry1 | Should -Be 'Hashtable'

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

        Add-WordTable -WordDocument $WordDocument -DataTable $InvoiceData -Design MediumShading1 -AutoFit Contents #-Verbose
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
