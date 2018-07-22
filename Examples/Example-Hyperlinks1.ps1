Import-Module PSWriteWord #-Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-HyperLinks1.docx"

$WordDocument = New-WordDocument -FilePath $FilePath

Add-WordText -WordDocument $WordDocument -Text 'This is my first title' -HeadingType Heading1 -Supress $True
$URL = Add-WordHyperLink -WordDocument $WordDocument -UrlText 'This is my url' -UrlLink 'https://evotec.xyz'
$Paragraph = Add-WordParagraph -WordDocument $WordDocument
Set-WordHyperLink -WordDocument $WordDocument -Paragraph $Paragraph -Value $URL -Supress $True

Save-WordDocument $WordDocument -Supress $True

Invoke-Item $FilePath