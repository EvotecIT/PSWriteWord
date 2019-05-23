Import-Module PSWriteWord -Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-ListItems1.docx"
$ListOfItemsNotArray1 = 'Test1'
$ListOfItemsNotArray2 = $false
$ListOfItemsNotArray3 = $false, $true

$WordDocument = New-WordDocument $FilePath

Add-WordText -WordDocument $WordDocument -Text 'This is text after which will be bulleted list' -FontSize 15 -UnderlineStyle singleLine -HeadingType Heading2 -Supress $True
Add-WordList -WordDocument $WordDocument -ListType Bulleted -ListData $ListOfItemsNotArray1 -Supress $True #-Verbose

Add-WordText -WordDocument $WordDocument -Text 'This is text after which will be bulleted list' -FontSize 15 -UnderlineStyle singleLine -HeadingType Heading2 -Supress $True
Add-WordList -WordDocument $WordDocument -ListType Bulleted -ListData $ListOfItemsNotArray2 -Supress $True #-Verbose

Add-WordText -WordDocument $WordDocument -Text 'This is text after which will be bulleted list' -FontSize 15 -UnderlineStyle singleLine -HeadingType Heading2 -Supress $True
Add-WordList -WordDocument $WordDocument -ListType Bulleted -ListData $ListOfItemsNotArray3 -Supress $True #-Verbose

Save-WordDocument $WordDocument -Language 'en-US' -Supress $true -OpenDocument