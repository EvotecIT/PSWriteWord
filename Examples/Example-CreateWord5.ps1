Import-Module PSWriteWord #-Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-CreateWord5.docx"

### define new document
$WordDocument = New-WordDocument $FilePath
### add 3 paragraphs, using pipeline for $worddocument
$WordDocument | Add-WordText -Text 'This is a text' -FontSize 10
$WordDocument | Add-WordText -Text 'This is a text font size 21' -FontSize 21
$WordDocument | Add-WordText -Text 'This is a text font size 15' -FontSize 15
### Save document
$WordDocument |Save-WordDocument

### Start Word with file
Invoke-Item $FilePath