Import-Module PSWriteWord #-Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-Equation.docx"

$WordDocument = New-WordDocument $FilePath
$Title1 = Add-WordText -WordDocument $WordDocument -Text 'This is an example showing ', 'how to add ', 'Equation to Microsoft Word' `
    -FontSize 10, 10, 10 `
    -Color Blue, Red, Blue `
    -Bold $false, $false, $true `
    -Italic $true, $true -SpacingAfter 10 -Supress $false

Set-WordParagraph -Paragraph $Title1 -Alignment center

Add-WordEquation -WordDocument $WordDocument -Equation "y = mx + b"

$Title2 = Add-WordText -WordDocument $WordDocument -Text 'This is 2nd example showing ', 'how to add ', 'Equation to Microsoft Word' `
    -FontSize 10, 10, 10 `
    -Color Blue, Red, Blue `
    -Bold $false, $false, $true `
    -Italic $true, $true -SpacingAfter 10 -Supress $false


Set-WordParagraph -Paragraph $Title2 -Alignment center

Add-WordEquation -WordDocument $WordDocument -Equation "x = ( -b (b - 4ac))/2a"

Save-WordDocument $WordDocument