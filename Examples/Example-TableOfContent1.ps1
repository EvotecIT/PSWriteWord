Import-Module PSWriteWord #-Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-TableOfContent1.docx"

$WordDocument = New-WordDocument -FilePath $FilePath

$toc = $WordDocument.InsertTableOfContents("Table of content", 1)

$p1 = $WordDocument.InsertParagraph("First")
$p1.StyleName = [HeadingType]::Heading1
$p1.Alignment = "left"
$p1.ListItemType = 'Numbered'

Add-Section -WordDocument $WordDocument -PageBreak

$p2 = $WordDocument.InsertParagraph("Second")
$p2.StyleName = [HeadingType]::Heading2

Add-Section -WordDocument $WordDocument -PageBreak

$p3 = $WordDocument.InsertParagraph("Third")
$p3.StyleName = [HeadingType]::Heading2

Save-WordDocument $WordDocument