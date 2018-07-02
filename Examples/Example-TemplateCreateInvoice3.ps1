Import-Module PSWriteWord #-Force

### Before running this script make sure to run Example-CreateWord first
$FilePathTemplate = "$PSScriptRoot\Templates\WordTemplate-InvoiceWithLogo.docx"
$FilePathInvoice = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-TemplateCreateInvoice3.docx"

$FilePathImage = "$PSScriptRoot\Images\Logo-Evotec-Small.jpg"

$WordDocument = Get-WordDocument -FilePath $FilePathTemplate

Add-WordCustomProperty -WordDocument $WordDocument -Name 'CompanyName'  -Value 'Evotec'
Add-WordCustomProperty -WordDocument $WordDocument -Name 'CompanySlogan'  -Value 'IT Consultants'
Add-WordCustomProperty -WordDocument $WordDocument -Name 'CompanyStreetName'  -Value 'Francuska 96B/23'
Add-WordCustomProperty -WordDocument $WordDocument -Name 'CompanyCity'  -Value 'Katowice'
Add-WordCustomProperty -WordDocument $WordDocument -Name 'CompanyZipCode'  -Value '40-507'
Add-WordCustomProperty -WordDocument $WordDocument -Name 'CompanyPhone'  -Value '+48 500 500 500'
Add-WordCustomProperty -WordDocument $WordDocument -Name 'CompanySupport'  -Value 'fake-email@evotec1.xyz'
Add-WordCustomProperty -WordDocument $WordDocument -Name 'ClientName'  -Value 'Fake Company'
Add-WordCustomProperty -WordDocument $WordDocument -Name 'ClientStreetName'  -Value 'Fake Street Name'
Add-WordCustomProperty -WordDocument $WordDocument -Name 'ClientCity'  -Value 'Warsaw'
Add-WordCustomProperty -WordDocument $WordDocument -Name 'ClientZipCode'  -Value '10-000'
Add-WordCustomProperty -WordDocument $WordDocument -Name 'ClientPhone'  -Value '+48 400 400 400'
Add-WordCustomProperty -WordDocument $WordDocument -Name 'ClientMail'  -Value 'fake-email@fake-company.com'

$ParagraphsWithPictures = Get-WordPicture -WordDocument $WordDocument -ListParagraphs
Set-WordPicture -WordDocument $WordDocument -Paragraph $ParagraphsWithPictures[0] -ImagePath $FilePathImage -ImageWidth 100 -ImageHeight 40

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

$LastTable = Get-WordTable -WordDocument $WordDocument -LastTable
$RowsToRemove = $LastTable.Rows.Count - 1
Remove-WordTableRow -Table $LastTable -Count $RowsToRemove
Add-WordTable -Table $LastTable -DataTable $InvoiceData -DoNotAddTitle -Verbose

Save-WordDocument -WordDocument $WordDocument -FilePath $FilePathInvoice
### Start Word with file
Invoke-Item $FilePathInvoice
