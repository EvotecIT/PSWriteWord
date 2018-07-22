Import-Module PSWriteWord #-Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-ListItems1.docx"
$ListOfItems = @('Test1', 'Test2', 'Test3', 'Test4', 'Test5')

$WordDocument = New-WordDocument $FilePath

Add-WordText -WordDocument $WordDocument -Text 'This is text after which will be bulleted list' -FontSize 15 -UnderlineStyle singleLine -HeadingType Heading2 -Supress $True
Add-WordList -WordDocument $WordDocument -ListType Bulleted -ListData $ListOfItems -Supress $True

Add-WordSection -WordDocument $WordDocument -PageBreak

Add-WordText -WordDocument $WordDocument -Text 'This is text after which will be numbered list' -FontSize 15 -UnderlineStyle singleLine -HeadingType Heading2 -Supress $True
Add-WordList -WordDocument $WordDocument -ListType Numbered -ListData $ListOfItems -Supress $True

Save-WordDocument $WordDocument -Language 'en-US' -Supress $true
Invoke-Item $FilePath