Import-Module PSWriteWord -Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-CreateWord2.docx"

$WordDocument = New-WordDocument $FilePath
$p1 = Add-WordText -WordDocument $WordDocument -Text 'This is a text aligned to center with Set-WordParagraph' -FontSize 10 -Spacingafter 50 -Supress $False
Set-WordParagraph -Paragraph $p1 -Alignment center

Add-WordParagraph -WordDocument $WordDocument # Adds an empty line

# Same action as above can be done with just one line.
$p1 = Add-WordText -WordDocument $WordDocument -Text 'This is a text aligned to center done with Add-WordText only' -FontSize 10 -Spacingafter 50 -Supress $False -Alignment center

Add-WordParagraph -WordDocument $WordDocument # Adds an empty line

$p1 = Add-WordText -WordDocument $WordDocument -Text 'This is a text to the left but with Right To Left' -FontSize 21 -Supress $false
Set-WordParagraph -Paragraph $p1 -Alignment left -Direction RightToLeft

Add-WordParagraph -WordDocument $WordDocument # Adds an empty line

$p2 = Add-WordText -WordDocument $WordDocument -Text 'This is a text that is justified.' -FontSize 15 -Supress $false
$p2 = Set-WordParagraph -Paragraph $p2 -Alignment Both -Direction LeftToRight -Supress $false
$p2 = Add-WordText -Paragraph $p2 -Text 'This text will append to last paragraph.' -FontSize 15 -Supress $false

Add-WordParagraph -WordDocument $WordDocument # Adds an empty line

Add-WordText -WordDocument $WordDocument -Text 'But you can actually just use one line to do Alingment and direction at same time' -FontSize 10 -Alignment Center -Direction LeftToRight

Save-WordDocument $WordDocument

### Start Word with file
Invoke-Item $FilePath