Import-Module PSWriteWord #-Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-PageOrientation2.docx"

$WordDocument = New-WordDocument $FilePath
### set orientation
Set-WordPageSettings -WordDocument $WordDocument -Orientation Landscape
### alternatively you can use this commandlet
Set-WordOrientation -WordDocument $WordDocument -Orientation Landscape
### add 3 paragraphs
Add-WordText -WordDocument $WordDocument -Text 'This is a text' -FontSize 10 -Supress $True
Add-WordText -WordDocument $WordDocument -Text 'This is a text font size 21' -FontSize 21 -Supress $True
Add-WordText -WordDocument $WordDocument -Text 'This is a text font size 15' -FontSize 15 -Supress $True
#Set-WordOrientation -WordDocument $WordDocument -Orientation Portrait

### get page settings
Get-WordPageSettings -WordDocument $WordDocument

# Add word section break
Add-WordSection -WordDocument $WordDocument -PageBreak

# Create new WORD Document
$WordDocument2 = New-WordDocument $FilePath
Set-WordPageSettings -WordDocument $WordDocument2 -Orientation Portrait
Add-WordText -WordDocument $WordDocument2 -Text 'This is a text' -FontSize 10 -Supress $True
Add-WordText -WordDocument $WordDocument2 -Text 'This is a text font size 21' -FontSize 21 -Supress $True


# Finally merge 2nd document into 1st document
$WordDocument.InsertDocument($WordDocument2)

### Save document
Save-WordDocument -WordDocument $WordDocument -Supress $True -OpenDocument