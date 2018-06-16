Import-Module PSWriteWord -Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-ListItems.docx"
$ListOfItems = @('Test1', 'Test2', 'Test3', 'Test4', 'Test5')

$WordDocument = New-WordDocument $FilePath

$p = $WordDocument.InsertParagraph("This is text after which will be bulleted list").FontSize(15)
Add-List -WordDocument $WordDocument -ListType Bulleted -ListData $ListOfItems

Add-Section -WordDocument $WordDocument -PageBreak

$p = $WordDocument.InsertParagraph("This is another text, after which will be numbered list").FontSize(15)
Add-List -WordDocument $WordDocument -ListType Numbered -ListData $ListOfItems

Save-WordDocument $WordDocument