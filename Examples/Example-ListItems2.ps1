Import-Module PSWriteWord -Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-ListItems2.docx"

$InvoiceEntry1 = @{}
$InvoiceEntry1.Description = 'IT Services'
$InvoiceEntry1.Amount = '$200'

$InvoiceEntry2 = @{}
$InvoiceEntry2.Description = 'IT Services'
$InvoiceEntry2.Amount = '$200'

$ListOfItems = @()
$ListOfItems += $InvoiceEntry1
$ListOfItems += $InvoiceEntry2

$WordDocument = New-WordDocument $FilePath

#Add-WordText -WordDocument $WordDocument -Text 'This is text after which will be bulleted list' -FontSize 15 -UnderlineStyle singleLine -HeadingType Heading2 -Supress $True -Verbose
Add-WordList -WordDocument $WordDocument -ListType Bulleted -ListData $ListOfItems -Supress $false -Verbose

#Add-WordSection -WordDocument $WordDocument -PageBreak -Supress $true

#Add-WordText -WordDocument $WordDocument -Text 'This is text after which will be numbered list' -FontSize 15 -UnderlineStyle singleLine -HeadingType Heading2 -Supress $True
#Add-WordList -WordDocument $WordDocument -ListType Numbered -ListData $ListOfItems -Supress $true

Save-WordDocument $WordDocument -Language 'en-US' -Supress $true
Invoke-Item $FilePath