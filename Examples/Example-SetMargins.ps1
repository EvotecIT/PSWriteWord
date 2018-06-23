Import-Module PSWriteWord -Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-SetMargins1.docx"

$WordDocument = New-WordDocument $FilePath
Add-WordText -WordDocument $WordDocument -Text 'This is a text' -FontSize 10
Add-WordText -WordDocument $WordDocument -Text 'This is a text font size 21' -FontSize 21
Add-WordText -WordDocument $WordDocument -Text 'This is a text font size 15' -FontSize 15

Set-WordMargins -WordDocument $WordDocument -MarginRight 85 -PageWidth 350
Save-WordDocument $WordDocument