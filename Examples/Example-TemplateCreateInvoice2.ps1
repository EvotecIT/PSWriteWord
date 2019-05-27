Import-Module PSWriteWord #-Force

### Before running this script make sure to run Example-CreateWord first
$FilePathTemplate = "$PSScriptRoot\Templates\WordTemplate-InvoiceWithLogo.docx"
$FilePathInvoice = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-TemplateCreateInvoice2.docx"

$FilePathImage = "$PSScriptRoot\Images\Logo-Evotec-Small.jpg"

$WordDocument = Get-WordDocument -FilePath $FilePathTemplate

Add-WordCustomProperty -WordDocument $WordDocument -Name 'CompanyName'  -Value 'Evotec' -Supress $True
Add-WordCustomProperty -WordDocument $WordDocument -Name 'CompanySlogan'  -Value 'IT Consultants' -Supress $True
Add-WordCustomProperty -WordDocument $WordDocument -Name 'CompanyStreetName'  -Value 'Francuska 96B/23' -Supress $True
Add-WordCustomProperty -WordDocument $WordDocument -Name 'CompanyCity'  -Value 'Katowice' -Supress $True
Add-WordCustomProperty -WordDocument $WordDocument -Name 'CompanyZipCode'  -Value '40-507' -Supress $True
Add-WordCustomProperty -WordDocument $WordDocument -Name 'CompanyPhone'  -Value '+48 500 500 500' -Supress $True
Add-WordCustomProperty -WordDocument $WordDocument -Name 'CompanySupport'  -Value 'fake-email@evotec1.xyz' -Supress $True
Add-WordCustomProperty -WordDocument $WordDocument -Name 'ClientName'  -Value 'Fake Company' -Supress $True
Add-WordCustomProperty -WordDocument $WordDocument -Name 'ClientStreetName'  -Value 'Fake Street Name' -Supress $True
Add-WordCustomProperty -WordDocument $WordDocument -Name 'ClientCity'  -Value 'Warsaw' -Supress $True
Add-WordCustomProperty -WordDocument $WordDocument -Name 'ClientZipCode'  -Value '10-000' -Supress $True
Add-WordCustomProperty -WordDocument $WordDocument -Name 'ClientPhone'  -Value '+48 400 400 400' -Supress $True
Add-WordCustomProperty -WordDocument $WordDocument -Name 'ClientMail'  -Value 'fake-email@fake-company.com' -Supress $True

$ParagraphsWithPictures = Get-WordPicture -WordDocument $WordDocument -ListParagraphs
# Since it's the only paragraph with picture.... we use 0
# Otherwise you would need to lopp thru ParagraphsWithPictures to find what you need
$ParagraphThatHasPicture = $ParagraphsWithPictures[0]
### You can do it manually... by removing picture and adding picture

#Remove-WordPicture -WordDocument $WordDocument -Paragraph $ParagraphsWithPictures[0]
#Add-WordPicture -WordDocument $WordDocument -Paragraph $ParagraphsWithPictures[0] -ImagePath $FilePathImage -ImageWidth 100 -ImageHeight 30

### Or you can do it as one command, keep in mind you can skip imagewidth/imageheight but I wanted to resize logo a bit
Set-WordPicture -WordDocument $WordDocument -Paragraph $ParagraphThatHasPicture -ImagePath $FilePathImage -ImageWidth 100 -ImageHeight 40

Save-WordDocument -WordDocument $WordDocument -FilePath $FilePathInvoice -Supress $True -OpenDocument
