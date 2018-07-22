Import-Module PSWriteWord -Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-ListItems2.docx"
$ListOfItems = @('Test1', 'Test2', 'Test3', 'Test4', 'Test5')

$WordDocument = New-WordDocument $FilePath

Add-WordText -WordDocument $WordDocument -Text 'This is text after which will be bulleted list' -FontSize 15 -UnderlineStyle singleLine -HeadingType Heading2 -Supress $True
Add-WordList -WordDocument $WordDocument -ListType Bulleted -ListData $ListOfItems -Supress $true -Verbose

Add-WordSection -WordDocument $WordDocument -PageBreak -Supress $true

Add-WordText -WordDocument $WordDocument -Text 'This is text after which will be numbered list' -FontSize 15 -UnderlineStyle singleLine -HeadingType Heading2 -Supress $True
Add-WordList -WordDocument $WordDocument -ListType Numbered -ListData $ListOfItems -Supress $true

Save-WordDocument $WordDocument -Language 'en-US' -Supress $true
Invoke-Item $FilePath