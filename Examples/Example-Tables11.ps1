
$InvoiceEntry7 = [ordered]@{}
$InvoiceEntry7.Description = 'IT Services 4'
$InvoiceEntry7.Amount = '$301'

$InvoiceEntry8 = [ordered]@{}
$InvoiceEntry8.Description = 'IT Services 5'
$InvoiceEntry8.Amount = '$299'

$InvoiceDataOrdered1 = @()
$InvoiceDataOrdered1 += $InvoiceEntry7

$InvoiceDataOrdered2 = @()
$InvoiceDataOrdered2 += $InvoiceEntry7
$InvoiceDataOrdered2 += $InvoiceEntry8



Import-Module PSWriteWord #-Force
$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-Tables11.docx"

$WordDocument = New-WordDocument $FilePath
$Table = Add-WordTable -WordDocument $WordDocument -DataTable $InvoiceDataOrdered1 -Design ColorfulGrid -Supress $false
$Table
Add-WordParagraph -WordDocument $WordDocument -Supress $True

Add-WordTable -WordDocument $WordDocument -DataTable $InvoiceDataOrdered2 -Design ColorfulGrid -Percentage $true

Save-WordDocument $WordDocument -Language 'en-US' -Supress $True
Invoke-Item $FilePath