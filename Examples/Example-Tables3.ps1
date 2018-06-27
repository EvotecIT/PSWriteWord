Import-Module PSWriteWord #-Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-Tables3.docx"

#Clear-Host
$WordDocument = New-WordDocument $FilePath

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

Add-WordText -WordDocument $WordDocument -Text "Invoice Data" -FontSize 15 -Alignment center
Add-WordParagraph -WordDocument $WordDocument
Add-WordTable -WordDocument $WordDocument -DataTable $InvoiceData -Design LightShading #-Verbose

Save-WordDocument $WordDocument

### Start Word with file
Invoke-Item $FilePath
