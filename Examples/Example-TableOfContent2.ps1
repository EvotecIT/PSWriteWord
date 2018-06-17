Import-Module PSWriteWord -Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-TableOfContent2.docx"

$WordDocument = New-WordDocument -FilePath $FilePath

$toc = $WordDocument.InsertTableOfContents("Table of content", 1)

Add-WordText -WordDocument $WordDocument -Text 'This is my first title' -HeadingType Heading1
Add-Section -WordDocument $WordDocument -PageBreak
Add-WordText -WordDocument $WordDocument -Text 'This is my second title' -HeadingType Heading1 -Color Red
Add-Section -WordDocument $WordDocument -PageBreak
Add-WordText  -WordDocument $WordDocument -Text 'This is my third title' -HeadingType Heading2
Save-WordDocument $WordDocument