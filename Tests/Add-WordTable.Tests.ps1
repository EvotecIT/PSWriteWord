#Requires -Modules Pester
Import-Module $PSScriptRoot\..\PSWriteWord.psd1 #-Force

Describe 'Add-WordTable' {
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

        $InvoiceData = @()
        $InvoiceData += $InvoiceEntry1
        $InvoiceData += $InvoiceEntry2
        $InvoiceData += $InvoiceEntry3
        $InvoiceData += $InvoiceEntry4
        $InvoiceData += $InvoiceEntry5

        Add-WordTable -WordDocument $WordDocument -DataTable $InvoiceData -Design MediumShading1 -AutoFit Contents #-Verbose
        $WordDocument.Tables[0].RowCount | Should -Be 6
        $WordDocument.Tables[0].ColumnCount | Should -Be 2
        #$WordDocument.Tables[0].AutoFit | Should -Be 'Contents' # Seems like a bug in Xceed - always returns ColumnWidth
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
