Import-Module PSWriteWord -Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-AddPicture1.docx"
$FilePathImage = "$PSScriptRoot\Images\Logo-Evotec-Small.jpg"

$WordDocument = New-WordDocument $FilePath

Add-WordText -WordDocument $WordDocument -Text 'Adding a picture...'

Add-WordPicture -WordDocument $WordDocument -ImagePath $FilePathImage -Verbose

Add-WordText -WordDocument $WordDocument -Text 'Adding a picture... with rotation'

Add-WordPicture -WordDocument $WordDocument -ImagePath $FilePathImage -Rotation 25

Add-WordText -WordDocument $WordDocument -Text 'Adding a picture... flip horizontal'

Add-WordPicture -WordDocument $WordDocument -ImagePath $FilePathImage -FlipHorizontal

Add-WordText -WordDocument $WordDocument -Text 'Adding a picture... flip horizontal and vertical'

Add-WordPicture -WordDocument $WordDocument -ImagePath $FilePathImage -FlipVertical -FlipHorizontal

Save-WordDocument $WordDocument -Language 'en-US'