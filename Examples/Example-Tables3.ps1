Import-Module PSWriteWord -Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-Tables3.docx"

#Clear-Host
$WordDocument = New-WordDocument $FilePath

$InvoiceEntry1 = @{}
$InvoiceEntry1.Description = 'IT Services'
$InvoiceEntry1.Amount = '$200'

$InvoiceEntry2 = @{}
$InvoiceEntry2.Description = 'IT Services'
$InvoiceEntry2.Amount = '$200'

$InvoiceData = @()
$InvoiceData += $InvoiceEntry1
$InvoiceData += $InvoiceEntry2

Add-WordText -WordDocument $WordDocument -Text "Invoice Data" -FontSize 15
Add-WordParagraph -WordDocument $WordDocument
Add-WordTable -WordDocument $WordDocument -Table $InvoiceData -Design LightShading  -Verbose

Save-WordDocument $WordDocument

### Start Word with file
Invoke-Item $FilePath
