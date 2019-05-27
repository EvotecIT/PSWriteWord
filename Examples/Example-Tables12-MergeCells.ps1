Import-Module pswriteword #-Force
$FilePath = "$Env:USERPROFILE\Desktop\Example-Tables12-MergeCells.docx"

$WordDocument = New-WordDocument $FilePath
$Table1 = New-WordTable -WordDocument $WordDocument -NrRows 3 -NrColumns 3
Set-WordTable -Table $Table1 -Design MediumGrid3Accent1 -Supress $true
Add-WordTableCellValue -Table $Table1 -Row 0 -Column 0 -Value "test1" -Supress $true
Add-WordTableCellValue -Table $Table1 -Row 0 -Column 0 -Value "test5" -Supress $true
Add-WordTableCellValue -Table $Table1 -Row 0 -Column 1 -Value "test2" -Supress $true
Add-WordTableCellValue -Table $Table1 -Row 0 -Column 1 -Value "test3" -Supress $true
Add-WordTableCellValue -Table $Table1 -Row 0 -Column 2 -Value "test4" -Supress $true
Set-WordTableRowMergeCells -Table $Table1 -RowNr 0 -ColumnNrStart 0 -ColumnNrEnd 2 -Supress $True -TextMerge
Save-WordDocument -WordDocument $WordDocument -OpenDocument -Supress $true