Import-Module $PSScriptRoot\..\PSWriteWord.psd1 -Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-AddPicture1.docx"
$FilePathImage = "$PSScriptRoot\Images\Logo-Evotec-Small.jpg"

$WordDocument = New-WordDocument $FilePath

Add-WordText -WordDocument $WordDocument -Text 'Adding a picture...' -Supress $true

$Picture = Add-WordPicture -WordDocument $WordDocument -ImagePath $FilePathImage -Verbose

#$Picture1 = Add-WordHyperLink -Paragraph $Picture -UrlText 'Test' -UrlLink 'https://evotec.xyz' -WordDocument $WordDocument

$Hyper = $WordDocument.AddHyperlink('f','https://evotec.xyz')
#$WordDocument.AddHyperlink

$Picture.AppendHyperlink($Hyper)
$Picture.InsertHyperlink(0,$Hyper)






#$Text = Add-WordText -WordDocument $WordDocument -Text 'Adding a picture... with rotation'
#Add-WordHyperLink -UrlLink 'https://evotec.xyz' -UrlText '' -WordDocument $WordDocument -Paragraph $Text
<#

Add-WordPicture -WordDocument $WordDocument -ImagePath $FilePathImage -Rotation 25 -Supress $true

Add-WordText -WordDocument $WordDocument -Text 'Adding a picture... flip horizontal' -Alignment right  -Supress $true

Add-WordPicture -WordDocument $WordDocument -ImagePath $FilePathImage -FlipHorizontal -Supress $true

Add-WordText -WordDocument $WordDocument -Text 'Adding a picture... flip horizontal and vertical'  -Supress $true

Add-WordPicture -WordDocument $WordDocument -ImagePath $FilePathImage -FlipVertical -FlipHorizontal -Supress $true
#>

Save-WordDocument $WordDocument -Language 'en-US' -Supress $true -OpenDocument