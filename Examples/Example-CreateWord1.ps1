Import-Module PSWriteWord #-Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-CreateWord1.docx"

### define new document
$WordDocument = New-WordDocument $FilePath
### add 3 paragraphs
Add-WordText -WordDocument $WordDocument -Text 'This is a text' -FontSize 10
Add-WordText -WordDocument $WordDocument -Text 'This is a text font size 21' -FontSize 21
Add-WordText -WordDocument $WordDocument -Text 'This is a text font size 15' -FontSize 15
### Save document
Save-WordDocument $WordDocument

### Start Word with file
Invoke-Item $FilePath