


Add-Section -WordDocument $WordDocument -PageBreak
$ListOfItems = @('Test1', 'Test2', 'Test3', 'Test4', 'Test5')
Add-List -WordDocument $WordDocument -ListType Bulleted -ListData $ListOfItems
$p = $WordDocument.InsertParagraph("This is another text").FontSize(15)
Add-List -WordDocument $WordDocument -ListType Numbered -ListData $ListOfItems
$p = $WordDocument.InsertParagraph("This is another text").FontSize(15)
