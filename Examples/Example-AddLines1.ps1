Import-Module PSWriteWord -Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-AddLines1.docx"

$WordDocument = New-WordDocument $FilePath
Add-WordText -WordDocument $WordDocument -Text 'This is a text' -FontSize 10

Add-WordLine -WordDocument $WordDocument -LineColor Red -LineType double

Add-WordText -WordDocument $WordDocument -Text 'This is a text font size 21' -FontSize 21

Add-WordLine -WordDocument $WordDocument -LineColor Blue -LineType single -LineSize 10

Add-WordText -WordDocument $WordDocument -Text 'This is a text font size 15' -FontSize 15

Save-WordDocument $WordDocument -Language 'en-US'