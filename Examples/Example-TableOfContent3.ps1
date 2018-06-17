Import-Module PSWriteWord -Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-TableOfContent3.docx"

$WordDocument = New-WordDocument -FilePath $FilePath

$toc = $WordDocument.InsertTableOfContents("Table of content", 1)

Add-WordText -WordDocument $WordDocument -Text 'This is my first title' -HeadingType Heading1
Add-Section -WordDocument $WordDocument -PageBreak
Add-WordText -WordDocument $WordDocument -Text 'This is my second title' -HeadingType Heading1 -Color Red -CapsStyle caps
Add-Section -WordDocument $WordDocument -PageBreak
Add-WordText  -WordDocument $WordDocument -Text 'This is my third title' -HeadingType Heading2 -Italic $true -Bold $true
Save-WordDocument $WordDocument