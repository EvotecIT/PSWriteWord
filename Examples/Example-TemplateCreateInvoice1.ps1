Import-Module PSWriteWord #-Force

### Before running this script make sure to run Example-CreateWord first
$FilePathTemplate = "$PSScriptRoot\Templates\WordTemplate-Invoice.docx"
$FilePathInvoice = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-TemplateCreateInvoice1.docx"

$WordDocument = Get-WordDocument -FilePath $FilePathTemplate
#$WordDocument

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

Save-WordDocument -WordDocument $WordDocument -FilePath $FilePathInvoice -Supress $True
### Start Word with file
Invoke-Item $FilePathInvoice
