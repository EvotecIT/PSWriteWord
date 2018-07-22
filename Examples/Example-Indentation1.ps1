Import-Module PSWriteWord -Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-Indentation1.docx"

$WordDocument = New-WordDocument $FilePath
Add-WordText -WordDocument $WordDocument -Text "Paragraph indentation" -FontSize 15 -Alignment center -SpacingAfter 50 -Supress $True
Add-WordText -WordDocument $WordDocument -Text "This is the first paragraph. It doesn't contain any indentation." -FontSize 10 -SpacingAfter 30 -Supress $True
Add-WordText -WordDocument $WordDocument -Text "This is the second paragraph. It contains an indentation on the first line." -FontSize 10 -IndentationFirstLine 1 -SpacingAfter 30 -Supress $True
Add-WordText -WordDocument $WordDocument -Text "This is the third paragraph. It contains an indentation on all the lines except the first one." -FontSize 10 -IndentationHanging 1 -SpacingAfter 30  -Supress $True

Set-WordPageSettings -WordDocument $WordDocument -PageWidth 250
Save-WordDocument -WordDocument $WordDocument -Language 'en-US' -Supress $True

Invoke-Item $FilePath