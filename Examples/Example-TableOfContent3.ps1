Import-Module PSWriteWord -Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-TableOfContent3.docx"

$WordDocument = New-WordDocument -FilePath $FilePath
$Toc = Add-WordTOC -WordDocument $WordDocument -Title 'Table of content' -HeaderStyle Heading2
Add-WordSection -WordDocument $WordDocument -PageBreak -Supress $True
Add-WordText -WordDocument $WordDocument -Text 'This is my first title' -HeadingType Heading1 -Supress $True
Add-WordSection -WordDocument $WordDocument -PageBreak -Supress $True
Add-WordText -WordDocument $WordDocument -Text 'This is my second title' -HeadingType Heading1 -Color Red -CapsStyle caps -Supress $True
Add-WordSection -WordDocument $WordDocument -PageBreak -Supress $True
Add-WordText  -WordDocument $WordDocument -Text 'This is my third title' -HeadingType Heading2 -Italic $true -Bold $true -Supress $True
Save-WordDocument $WordDocument -Supress $True
### Start Word with file
Invoke-Item $FilePath