$FilePath = "$Env:USERPROFILE\Desktop\Example-Tables12-MergeCells-3.docx"

$WordDocument = New-WordDocument $FilePath
$Table1 = New-WordTable -WordDocument $WordDocument -NrRows 3 -NrColumns 3
Set-WordTable -Table $Table1 -Design MediumGrid3Accent1 -Supress $true
Add-WordTableCellValue -Table $Table1 -Row 0 -Column 0 -Value "test00" -Supress $true
Add-WordTableCellValue -Table $Table1 -Row 0 -Column 1 -Value "test01" -Supress $true
Add-WordTableCellValue -Table $Table1 -Row 0 -Column 2 -Value "test02" -Supress $true
$Cell = Add-WordTableCellValue -Table $Table1 -Row 1 -Column 0 -Value "test10"
$Cell = Add-WordText -WordDocument $WordDocument -Paragraph $Cell -Text "Test101"
$Cell = Add-WordText -WordDocument $WordDocument -Paragraph $Cell -Text "Test102"
$Cell = Add-WordText -WordDocument $WordDocument -Paragraph $Cell -Text "Test103"
$Cell = Add-WordText -WordDocument $WordDocument -Paragraph $Cell -Text "Test104"
Add-WordTableCellValue -Table $Table1 -Row 1 -Column 1 -Value "test11" -Supress $true
Add-WordTableCellValue -Table $Table1 -Row 2 -Column 0 -Value "test20" -Supress $true
$Cell = Add-WordTableCellValue -Table $Table1 -Row 2 -Column 1 -Value "test21"
Add-WordText -WordDocument $WordDocument -Paragraph $Cell -Text "Test201" -Supress $true
Add-WordTableCellValue -Table $Table1 -Row 2 -Column 2 -Value "test22" -Supress $true

Set-WordTableRowMergeCells -Table $Table1 -RowNr 1 -ColumnNrStart 0 -ColumnNrEnd 2 -Supress $True -TextMerge
Save-WordDocument -WordDocument $WordDocument -OpenDocument -Supress $true