Import-Module PSWriteWord #-Force
$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-TableOfContent7.docx"

$WordDocument = New-WordDocument -FilePath $FilePath
$WordDocument | Add-WordToc -Title 'Table of content' -Switches C, A -RightTabPos 15 -HeaderStyle Heading1 -Supress $true

$Section = "Test 1"

$Paragraph = $WordDocument | Add-WordTocItem -Text $Section -ListLevel 0 -ListItemType Numbered -HeadingType Heading1
$Paragraph = $WordDocument | Add-WordText -Text 'This is my test. Added after TOC Item.' -Color Orange

$Paragraph = $WordDocument | Add-WordTocItem -Text $Section -ListLevel 1 -ListItemType Numbered -HeadingType Heading1
$Paragraph = $WordDocument | Add-WordText -Text 'This is my test. Added after TOC Item.' -Color Orange

$Paragraph = $WordDocument | Add-WordTocItem -Text $Section -ListLevel 2 -ListItemType Numbered -HeadingType Heading1
$Paragraph = $WordDocument | Add-WordText -Text 'This is my test. Added after TOC Item.' -Color Orange

$Paragraph = $WordDocument | Add-WordTocItem -Text $Section -ListLevel 0 -ListItemType Numbered -HeadingType Heading2
$Paragraph = $WordDocument | Add-WordText -Text 'This is my test. Added after TOC Item.' -Color Orange

$Paragraph = $WordDocument | Add-WordTocItem -Text 'Adding another one' -ListLevel 3 -ListItemType Numbered -HeadingType Heading3
$Paragraph = $WordDocument | Add-WordText -Text 'This is my test - outside of loop. Added after TOC Item.' -Color Red

$WordDocument | Save-WordDocument -Language 'en-US' -Supress $True
### Start Word with file
Invoke-Item $FilePath