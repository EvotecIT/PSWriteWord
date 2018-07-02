Import-Module PSWriteWord #-Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-CreateWord2.docx"

$WordDocument = New-WordDocument $FilePath
$p1 = Add-WordText -WordDocument $WordDocument -Text 'This is a text aligned to centerr' -FontSize 10 -Spacingafter 50 -Supress $False
Set-WordParagraph -Paragraph $p1 -Alignment center

$p1 = Add-WordText -WordDocument $WordDocument -Text 'This is a text to the left but with Right To Left' -FontSize 21 -Supress $false
Set-WordParagraph -Paragraph $p1 -Alignment left -Direction RightToLeft

$p1 = Add-WordText -WordDocument $WordDocument -Text 'This is a text that is justified.' -FontSize 15 -Supress $false
Set-WordParagraph -Paragraph $p1 -Alignment Both -Direction LeftToRight

Add-WordText -WordDocument $WordDocument -Text 'But you can actually just use one line to do Alingment and direction at same time' -FontSize 10 -Alignment Center -Direction LeftToRight

Save-WordDocument $WordDocument

### Start Word with file
Invoke-Item $FilePath