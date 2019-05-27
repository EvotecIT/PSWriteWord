Import-Module PSWriteWord #-Force

### Before running this script make sure to run Example-CreateWord first
$FilePathTemplate = "$PSScriptRoot\Templates\WordTemplate-InvoiceWithLogo.docx"
$FilePathInvoice = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-TemplateCreateInvoice3.docx"

$FilePathImage = "$PSScriptRoot\Images\Logo-Evotec-Small.jpg"

$WordDocument = Get-WordDocument -FilePath $FilePathTemplate

Add-WordCustomProperty -WordDocument $WordDocument -Name 'CompanyName'  -Value 'Evotec'  -Supress $true
Add-WordCustomProperty -WordDocument $WordDocument -Name 'CompanySlogan'  -Value 'IT Consultants'  -Supress $true
Add-WordCustomProperty -WordDocument $WordDocument -Name 'CompanyStreetName'  -Value 'Francuska 96B/23'  -Supress $true
Add-WordCustomProperty -WordDocument $WordDocument -Name 'CompanyCity'  -Value 'Katowice' -Supress $true
Add-WordCustomProperty -WordDocument $WordDocument -Name 'CompanyZipCode'  -Value '40-507' -Supress $true
Add-WordCustomProperty -WordDocument $WordDocument -Name 'CompanyPhone'  -Value '+48 500 500 500' -Supress $true
Add-WordCustomProperty -WordDocument $WordDocument -Name 'CompanySupport'  -Value 'fake-email@evotec1.xyz' -Supress $true
Add-WordCustomProperty -WordDocument $WordDocument -Name 'ClientName'  -Value 'Fake Company' -Supress $true
Add-WordCustomProperty -WordDocument $WordDocument -Name 'ClientStreetName'  -Value 'Fake Street Name'-Supress $true
Add-WordCustomProperty -WordDocument $WordDocument -Name 'ClientCity'  -Value 'Warsaw' -Supress $true
Add-WordCustomProperty -WordDocument $WordDocument -Name 'ClientZipCode'  -Value '10-000' -Supress $true
Add-WordCustomProperty -WordDocument $WordDocument -Name 'ClientPhone'  -Value '+48 400 400 400' -Supress $true
Add-WordCustomProperty -WordDocument $WordDocument -Name 'ClientMail'  -Value 'fake-email@fake-company.com' -Supress $true

$ParagraphsWithPictures = Get-WordPicture -WordDocument $WordDocument -ListParagraphs
Set-WordPicture -WordDocument $WordDocument -Paragraph $ParagraphsWithPictures[0] -ImagePath $FilePathImage -ImageWidth 100 -ImageHeight 40 -Supress $true

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

$InvoiceData = @(
    $InvoiceEntry1
    $InvoiceEntry2
    $InvoiceEntry3
    $InvoiceEntry4
    $InvoiceEntry5
)

# Edit table, first find it
$LastTable = Get-WordTable -WordDocument $WordDocument -LastTable

# Remove last row
$RowsToRemove = $LastTable.Rows.Count - 1
Remove-WordTableRow -Table $LastTable -Count $RowsToRemove -Supress $true

# add new table
Add-WordTable -Table $LastTable -DataTable $InvoiceData -DoNotAddTitle -Supress $true


Save-WordDocument -WordDocument $WordDocument -FilePath $FilePathInvoice -Supress $true -OpenDocument
