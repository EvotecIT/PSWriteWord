Import-Module PSWriteWord #-Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-Tables7.docx"

#Clear-Host
$WordDocument = New-WordDocument $FilePath

$InvoiceEntry1 = @{ Description = 'IT Services 1'; Amount = '$230' }
$InvoiceEntry2 = @{ Description = 'IT Services 2'; Amount = '$200' }

$InvoiceData = @()
$InvoiceData += $InvoiceEntry1
$InvoiceData += $InvoiceEntry2

$InvoiceData1 = @()
$InvoiceData1 += $InvoiceEntry1

Add-WordText -WordDocument $WordDocument -Text "Invoice Data with 2 entries" -FontSize 15 -Alignment center -Supress $true
Add-WordParagraph -WordDocument $WordDocument -Supress $true
Add-WordTable -WordDocument $WordDocument -DataTable $InvoiceData -Design LightShading -Supress $true -Transpose #-Verbose

Add-WordText -WordDocument $WordDocument -Text "Invoice Data with just 1 entry" -FontSize 15 -Alignment center -Supress $true
Add-WordParagraph -WordDocument $WordDocument -Supress $true
Add-WordTable -WordDocument $WordDocument -DataTable $InvoiceData1 -Design LightShading -Supress $true -Transpose#-Verbose

Add-WordText -WordDocument $WordDocument -Text "Invoice Data with 2 entries - autofit to contents" -FontSize 15 -Alignment center -Supress $true
Add-WordParagraph -WordDocument $WordDocument -Supress $true
Add-WordTable -WordDocument $WordDocument -DataTable $InvoiceData -Design LightShading -AutoFit Contents -Supress $true -Transpose #-Verbose

Add-WordText -WordDocument $WordDocument -Text "Invoice Data with just 1 entry - autofit to window" -FontSize 15 -Alignment center -Supress $true
Add-WordParagraph -WordDocument $WordDocument -Supress $true
Add-WordTable -WordDocument $WordDocument -DataTable $InvoiceData1 -Design LightShading -AutoFit Window -Supress $true -Transpose

Add-WordText -WordDocument $WordDocument -Text "Invoice Data with 2 entries - direction right to left" -FontSize 15 -Alignment center -Supress $true
Add-WordParagraph -WordDocument $WordDocument -Supress $true
Add-WordTable -WordDocument $WordDocument -DataTable $InvoiceData -Design LightShading -Direction RightToLeft -Supress $true -Transpose #-Verbose

Add-WordText -WordDocument $WordDocument -Text "Invoice Data with just 1 entry - with break page before table" `
    -FontSize 15 -Alignment center -ShadingType Paragraph -ShadingColor Blue -Supress $true
Add-WordParagraph -WordDocument $WordDocument -Supress $true
Add-WordTable -WordDocument $WordDocument -DataTable $InvoiceData1 -Design LightShading -BreakPageBeforeTable -Supress $true -Transpose

Save-WordDocument $WordDocument -Language 'en-US' -Supress $true

### Start Word with file
Invoke-Item $FilePath