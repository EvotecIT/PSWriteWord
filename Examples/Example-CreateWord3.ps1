Import-Module PSWriteWord #-Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-CreateWord3.docx"

$WordDocument = New-WordDocument $FilePath
Add-WordText -WordDocument $WordDocument -Text 'This is a text' -FontSize 10 -Supress $True
Add-WordText -WordDocument $WordDocument -Text 'This is a text' -FontSize 10 -Supress $True
Add-WordText -WordDocument $WordDocument -Text 'This is a text with Heading type 3' -FontSize 10 -HeadingType Heading3 -Supress $True

Save-WordDocument $WordDocument -Supress $True

### Start Word with file
Invoke-Item $FilePath