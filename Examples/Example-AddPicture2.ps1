Import-Module PSWriteWord #-Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-AddPicture2.docx"
$FilePathImage1 = "$PSScriptRoot\Images\Logo-Evotec-Small.png"
$FilePathImage2 = "$PSScriptRoot\Images\Logo-Evotec-Small.jpg"

$WordDocument = New-WordDocument $FilePath

Add-WordText -WordDocument $WordDocument -Text 'Adding a picture...' -Supress $true
Add-WordPicture -WordDocument $WordDocument -ImagePath $FilePathImage1 -Verbose
Add-WordText -WordDocument $WordDocument -Text 'Adding a picture... with rotation' -Supress $true
Add-WordPicture -WordDocument $WordDocument -ImagePath $FilePathImage2 -Rotation 25

$PlaceToAddPicture = Add-WordText -WordDocument $WordDocument -Text 'Adding a picture...' -Supress $false

Add-WordText -WordDocument $WordDocument -Text 'This is text' -Supress $true

$AllPictures = Get-WordPicture -WordDocument $WordDocument -ListPictures
Add-WordText -WordDocument $WordDocument -Text 'This is another text' -Supress $true
Add-WordPicture -WordDocument $WordDocument -Picture $AllPictures[1] # add copy of picture
Add-WordPicture -WordDocument $WordDocument -Picture $AllPictures[1] -Paragraph $PlaceToAddPicture # add copy of picture to paragraph

Add-WordText -WordDocument $WordDocument -Text 'Here we copy 1st picture from WordDocument and add it again'  -Supress $true
$Picture = Get-WordPicture -WordDocument $WordDocument -PictureID 0
Add-WordPicture -WordDocument $WordDocument -Picture $Picture # add copy of picture

Save-WordDocument $WordDocument -Language 'en-US' -Supress $true

### Start Word with file
Invoke-Item $FilePath