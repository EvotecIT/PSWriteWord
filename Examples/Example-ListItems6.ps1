Import-Module PSWriteWord -Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-ListItems6.docx"
$ListOfItems = @('Test1', 'Test2')

$WordDocument = New-WordDocument $FilePath
Add-WordParagraph -WordDocument $WordDocument
$List = Add-WordList -WordDocument $WordDocument -ListType Numbered -ListData $ListOfItems -Supress $false
Set-WordList -List $List -FontSize 6 -FontFamily 'Tahoma' -Color Orange -Supress $True

$WordDocument.Paragraphs
$WordDocument.Lists
$WordDocument.ParagraphsDeepSearch


Save-WordDocument $WordDocument -Language 'en-US' -Supress $true
Invoke-Item $FilePath