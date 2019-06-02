Import-Module PSWriteWord #-Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-PageOrientation1.docx"

$WordDocument = New-WordDocument $FilePath
### set orientation
Set-WordPageSettings -WordDocument $WordDocument -Orientation Landscape

### alternatively you can use this commandlet
Set-WordOrientation -WordDocument $WordDocument -Orientation Landscape

### add 3 paragraphs
Add-WordText -WordDocument $WordDocument -Text 'This is a text' -FontSize 10 -Supress $True
Add-WordText -WordDocument $WordDocument -Text 'This is a text font size 21' -FontSize 21 -Supress $True
Add-WordText -WordDocument $WordDocument -Text 'This is a text font size 15' -FontSize 15 -Supress $True

### get page settings
Get-WordPageSettings -WordDocument $WordDocument
### Save document
Save-WordDocument -WordDocument $WordDocument -Supress $True -OpenDocument