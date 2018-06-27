Import-Module PSWriteWord #-Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-Tables7.docx"

#Clear-Host
$WordDocument = New-WordDocument $FilePath

$InvoiceEntry1 = @{ Description = 'IT Services 1'; Amount = '$230' }
$InvoiceEntry2 = @{ Description = 'IT Services 2'; Amount = '$200' }

$InvoiceData = @()
$InvoiceData += $InvoiceEntry1
$InvoiceData += $InvoiceEntry2

Add-WordText -WordDocument $WordDocument -Text "Invoice Data" -FontSize 15
Add-WordParagraph -WordDocument $WordDocument
Add-WordTable -WordDocument $WordDocument -DataTable $InvoiceData -Design LightShading #-Verbose

Save-WordDocument $WordDocument -Language 'en-US'

### Start Word with file
Invoke-Item $FilePath