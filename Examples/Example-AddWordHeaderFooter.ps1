Import-Module ..\PSWriteWord -Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-AddWordHeaderFooter.docx"

### define new document
$WordDocument = New-WordDocument $FilePath -Verbose

$Footer = Add-WordFooter -WordDocument $WordDocument -DifferentFirstPage $true -DifferentOddAndEvenPages $false
$Header = Add-WordHeader -WordDocument $WordDocument
### add 3 paragraphs
Add-WordText -WordDocument $WordDocument -Text 'This is a text' -FontSize 10 -Supress $True
$Paragraph = Add-WordPageBreak -WordDocument $WordDocument -Verbose
Add-WordText -WordDocument $WordDocument -Text 'This is a text font size 21' -FontSize 21 -Supress $True
$Paragraph = Add-WordText -WordDocument $WordDocument -Text 'This is a text font size 15' -FontSize 15 -Supress $false
$Paragraph | Add-WordPageBreak -InsertWhere BeforeSelf -Supress $True -Verbose

# this appends text to paragraph that already exists within Footer (when footer is created, first paragraph is also created)
Add-WordText -WordDocument $WordDocument -Footer $Footer.First -Paragraph $Footer.First.Paragraphs[0] -AppendToExistingParagraph -Text 'My Text in Footer - 1st paragraph' -Color Orange -Supress $True

Add-WordText -WordDocument $WordDocument -Footer $Footer.First -Text 'My Text in Footer - Paragraph will be added' -Color Red -Supress $True

Add-WordText -WordDocument $WordDocument -Header $Header.First -Text 'My Text in Header' -Color Blue -Supress $True

### Save document
Save-WordDocument $WordDocument -Supress $true -Language 'en-US' -Verbose -OpenDocument