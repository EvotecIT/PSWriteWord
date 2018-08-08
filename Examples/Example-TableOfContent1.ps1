Import-Module PSWriteWord -Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-TableOfContent1.docx"

$WordDocument = New-WordDocument -FilePath $FilePath

$Toc = Add-WordTOC -WordDocument $WordDocument -Title 'Table of content'
Add-WordText -WordDocument $WordDocument -HeadingType Heading1 -Text 'First' -Supress $True
Add-WordSection -WordDocument $WordDocument -PageBreak -Supress $True
Add-WordText -WordDocument $WordDocument -HeadingType Heading2 -Text 'Second' -Supress $True
Add-WordSection -WordDocument $WordDocument -PageBreak -Supress $True
Add-WordText -WordDocument $WordDocument -HeadingType Heading1 -Text 'Third' -Supress $True
Save-WordDocument $WordDocument -Supress $True
Invoke-Item $FilePath