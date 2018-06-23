
Import-Module PSWriteWord -Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-ParagraphAdvanced2.docx"

$WordDocument = New-WordDocument $FilePath
Add-WordText -WordDocument $WordDocument -Text 'This is text that has font size of 15', ' and this is font size of 11 ', ' while this will be 12.' `
    -FontSize 15, 10 `
    -Color Blue, Red `
    -Bold $true, $false, $true `
    -Italic $true, $true  -Language 'en-US', 'en-US', 'en-US'


Add-WordText -WordDocument $WordDocument -Text 'This is text that has font size of 15', ' and this is font size of 10 ', ' while this will be 12.' `
    -FontSize 15, 10 `
    -Color Blue, Red `
    -Bold $true, $false, $true `
    -Italic $true, $true `
    -FontName 'Tahoma', 'Arial', 'Times New Roman' `
    -SpacingAfter 5, 8 `
    -StrikeThrough doubleStrike, strike `
    -Highlight darkCyan `
    -SpacingBefore 15, 50 `
    -Spacing 5, 20, 1 `
    -NewLine $true, $true  -Language 'en-US', 'en-US', 'en-US'


Add-WordText -WordDocument $WordDocument -Text 'This is text that has font size of 15', ' and this is font size of 10 ', ' while this will be default size.' `
    -FontSize 15, 10 `
    -Color Blue, Red `
    -NewLine $true, $true -Language 'en-US', 'en-US', 'en-US'


Add-WordText -WordDocument $WordDocument -Text 'To jest polska wersja tekstu.' `
    -FontSize 15, 10 `
    -Color Blue, Red `
    -NewLine $true, $true -Language 'pl-PL'


Save-WordDocument -WordDocument $WordDocument