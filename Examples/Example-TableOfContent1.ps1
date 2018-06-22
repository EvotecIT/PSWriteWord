Import-Module PSWriteWord #-Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-TableOfContent1.docx"

$WordDocument = New-WordDocument -FilePath $FilePath

$Toc = Add-WordTOC -WordDocument $WordDocument -Title 'Table of content' -Switches S
Add-WordText -WordDocument $WordDocument -HeadingType Heading1 -Text 'First'
Add-WordSection -WordDocument $WordDocument -PageBreak
Add-WordText -WordDocument $WordDocument -HeadingType Heading2 -Text 'Second'
Add-WordSection -WordDocument $WordDocument -PageBreak
Add-WordText -WordDocument $WordDocument -HeadingType Heading1 -Text 'Third'
Save-WordDocument $WordDocument