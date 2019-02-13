Describe 'Set-WordTableRowMergeCells - Should Merge 3 columns and keep text from 3 columns merged' {
    It 'Add 4 values in 3 top columns and merge them, keep text' {
        $WordDocument = New-WordDocument
        $Table1 = New-WordTable -WordDocument $WordDocument -NrRows 3 -NrColumns 3

        $Table1.ColumnCount | Should -Be 3
        $Table1.RowCount | Should -Be 3

        Set-WordTable -Table $Table1 -Design MediumGrid3Accent1 -Supress $true

        $Table1.Design | Should -Be 'MediumGrid3Accent1'

        Add-WordTableCellValue -Table $Table1 -Row 0 -Column 0 -Value "test1" -Supress $true
        Add-WordTableCellValue -Table $Table1 -Row 0 -Column 1 -Value "test2" -Supress $true
        Add-WordTableCellValue -Table $Table1 -Row 0 -Column 1 -Value "test3" -Supress $true
        Add-WordTableCellValue -Table $Table1 -Row 0 -Column 2 -Value "test4" -Supress $true

        $Table1.Rows[0].Cells[0].Paragraphs[0].Text | Should -Be 'test1'
        $Table1.Rows[0].Cells[1].Paragraphs[0].Text | Should -Be 'test2test3'
        $Table1.Rows[0].Cells[2].Paragraphs[0].Text | Should -Be 'test4'

        Set-WordTableRowMergeCells -Table $Table1 -RowNr 0 -ColumnNrStart 0 -ColumnNrEnd 2 -Supress $True -TextMerge

        $Table1.Rows[0].Cells[0].Paragraphs[0].Text | Should -Be 'test1 test2test3 test4'
    }
    It 'Add 3 values in 3 top columns and merge them, keep text' {
        $WordDocument = New-WordDocument
        $Table1 = New-WordTable -WordDocument $WordDocument -NrRows 3 -NrColumns 3

        $Table1.ColumnCount | Should -Be 3
        $Table1.RowCount | Should -Be 3

        Set-WordTable -Table $Table1 -Design MediumGrid3Accent1 -Supress $true

        $Table1.Design | Should -Be 'MediumGrid3Accent1'

        Add-WordTableCellValue -Table $Table1 -Row 0 -Column 0 -Value "test1" -Supress $true
        Add-WordTableCellValue -Table $Table1 -Row 0 -Column 1 -Value "test2" -Supress $true
        Add-WordTableCellValue -Table $Table1 -Row 0 -Column 1 -Value "test3" -Supress $true

        $Table1.Rows[0].Cells[0].Paragraphs[0].Text | Should -Be 'test1'
        $Table1.Rows[0].Cells[1].Paragraphs[0].Text | Should -Be 'test2test3'

        Set-WordTableRowMergeCells -Table $Table1 -RowNr 0 -ColumnNrStart 0 -ColumnNrEnd 2 -Supress $True -TextMerge

        $Table1.Rows[0].Cells[0].Paragraphs[0].Text | Should -Be 'test1 test2test3'
    }
    It 'Add 0 values in 3 top columns and merge them, use merge text, but nothing to merge' {
        $WordDocument = New-WordDocument
        $Table1 = New-WordTable -WordDocument $WordDocument -NrRows 3 -NrColumns 3

        $Table1.ColumnCount | Should -Be 3
        $Table1.RowCount | Should -Be 3

        Set-WordTable -Table $Table1 -Design MediumGrid3Accent1 -Supress $true

        $Table1.Design | Should -Be 'MediumGrid3Accent1'

        $Table1.Rows[0].Cells[0].Paragraphs[0].Text | Should -Be ''
        $Table1.Rows[0].Cells[1].Paragraphs[0].Text | Should -Be ''
        $Table1.Rows[0].Cells[2].Paragraphs[0].Text | Should -Be ''

        Set-WordTableRowMergeCells -Table $Table1 -RowNr 0 -ColumnNrStart 0 -ColumnNrEnd 2 -Supress $True -TextMerge

        $Table1.Rows[0].Cells[0].Paragraphs[0].Text | Should -Be ''
    }
    It 'Add 4 values in 3 top columns and merge them and keep text only on 1st column, then add 3 values in 1st row in different columns, merge them and merge text' {
        $WordDocument = New-WordDocument
        $Table1 = New-WordTable -WordDocument $WordDocument -NrRows 3 -NrColumns 3

        $Table1.ColumnCount | Should -Be 3
        $Table1.RowCount | Should -Be 3

        Set-WordTable -Table $Table1 -Design MediumGrid3Accent1 -Supress $true

        $Table1.Design | Should -Be 'MediumGrid3Accent1'

        Add-WordTableCellValue -Table $Table1 -Row 0 -Column 0 -Value "test1" -Supress $true
        Add-WordTableCellValue -Table $Table1 -Row 0 -Column 1 -Value "test2" -Supress $true
        Add-WordTableCellValue -Table $Table1 -Row 0 -Column 1 -Value "test3" -Supress $true
        Add-WordTableCellValue -Table $Table1 -Row 0 -Column 2 -Value "test4" -Supress $true

        $Cell = Add-WordTableCellValue -Table $Table1 -Row 1 -Column 0 -Value "test10" 
        Add-WordText -WordDocument $WordDocument -Paragraph $Cell -Text "Test101" -Supress $true 
        Add-WordTableCellValue -Table $Table1 -Row 1 -Column 1 -Value "test11" -Supress $true
        Add-WordTableCellValue -Table $Table1 -Row 2 -Column 0 -Value "test20" -Supress $true

        $Table1.Rows[0].Cells[0].Paragraphs[0].Text | Should -Be 'test1'
        $Table1.Rows[0].Cells[1].Paragraphs[0].Text | Should -Be 'test2test3'
        $Table1.Rows[0].Cells[2].Paragraphs[0].Text | Should -Be 'test4'

        Set-WordTableRowMergeCells -Table $Table1 -RowNr 0 -ColumnNrStart 0 -ColumnNrEnd 2 -Supress $True
        Set-WordTableRowMergeCells -Table $Table1 -RowNr 1 -ColumnNrStart 0 -ColumnNrEnd 2 -Supress $True -TextMerge


        $Table1.Rows[0].Cells[0].Paragraphs[0].Text | Should -Be 'test1'
        $Table1.Rows[1].Cells[0].Paragraphs[0].Text | Should -Be 'test10'
        $Table1.Rows[1].Cells[0].Paragraphs[1].Text | Should -Be 'test101 test11'
    }

}

Describe 'Set-WordTableRowMergeCells - Should Merge 3 columns and keep text from 1st column only' {
    It 'Add 4 values in 3 top columns and merge them, keep text only in 1st column' {
        $WordDocument = New-WordDocument
        $Table1 = New-WordTable -WordDocument $WordDocument -NrRows 3 -NrColumns 3

        $Table1.ColumnCount | Should -Be 3
        $Table1.RowCount | Should -Be 3

        Set-WordTable -Table $Table1 -Design MediumGrid3Accent1 -Supress $true

        $Table1.Design | Should -Be 'MediumGrid3Accent1'

        Add-WordTableCellValue -Table $Table1 -Row 0 -Column 0 -Value "test1" -Supress $true
        Add-WordTableCellValue -Table $Table1 -Row 0 -Column 1 -Value "test2" -Supress $true
        Add-WordTableCellValue -Table $Table1 -Row 0 -Column 1 -Value "test3" -Supress $true
        Add-WordTableCellValue -Table $Table1 -Row 0 -Column 2 -Value "test4" -Supress $true

        $Table1.Rows[0].Cells[0].Paragraphs[0].Text | Should -Be 'test1'
        $Table1.Rows[0].Cells[1].Paragraphs[0].Text | Should -Be 'test2test3'
        $Table1.Rows[0].Cells[2].Paragraphs[0].Text | Should -Be 'test4'

        Set-WordTableRowMergeCells -Table $Table1 -RowNr 0 -ColumnNrStart 0 -ColumnNrEnd 2 -Supress $True

        $Table1.Rows[0].Cells[0].Paragraphs[0].Text | Should -Be 'test1'
    }
    It 'Add 3 values in 3 top columns and merge them, keep text only in 1st column' {
        $WordDocument = New-WordDocument
        $Table1 = New-WordTable -WordDocument $WordDocument -NrRows 3 -NrColumns 3

        $Table1.ColumnCount | Should -Be 3
        $Table1.RowCount | Should -Be 3

        Set-WordTable -Table $Table1 -Design MediumGrid3Accent1 -Supress $true

        $Table1.Design | Should -Be 'MediumGrid3Accent1'

        Add-WordTableCellValue -Table $Table1 -Row 0 -Column 0 -Value "test1" -Supress $true
        Add-WordTableCellValue -Table $Table1 -Row 0 -Column 1 -Value "test2" -Supress $true
        Add-WordTableCellValue -Table $Table1 -Row 0 -Column 1 -Value "test3" -Supress $true

        $Table1.Rows[0].Cells[0].Paragraphs[0].Text | Should -Be 'test1'
        $Table1.Rows[0].Cells[1].Paragraphs[0].Text | Should -Be 'test2test3'

        Set-WordTableRowMergeCells -Table $Table1 -RowNr 0 -ColumnNrStart 0 -ColumnNrEnd 2 -Supress $True

        $Table1.Rows[0].Cells[0].Paragraphs[0].Text | Should -Be 'test1'
    }
    It 'Add 0 values in 3 top columns and merge them, use merge text, but nothing to merge' {
        $WordDocument = New-WordDocument
        $Table1 = New-WordTable -WordDocument $WordDocument -NrRows 3 -NrColumns 3

        $Table1.ColumnCount | Should -Be 3
        $Table1.RowCount | Should -Be 3

        Set-WordTable -Table $Table1 -Design MediumGrid3Accent1 -Supress $true

        $Table1.Design | Should -Be 'MediumGrid3Accent1'

        $Table1.Rows[0].Cells[0].Paragraphs[0].Text | Should -Be ''
        $Table1.Rows[0].Cells[1].Paragraphs[0].Text | Should -Be ''
        $Table1.Rows[0].Cells[2].Paragraphs[0].Text | Should -Be ''

        Set-WordTableRowMergeCells -Table $Table1 -RowNr 0 -ColumnNrStart 0 -ColumnNrEnd 2 -Supress $True

        $Table1.Rows[0].Cells[0].Paragraphs[0].Text | Should -Be ''
    }

}
