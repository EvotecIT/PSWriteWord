Import-Module .\PSWriteWord.psd1 -Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-HyperLinks2.docx"

$WordDocument = New-WordDocument -FilePath $FilePath

$Paragraph = Add-WordText -WordDocument $WordDocument -Text 'This is my first title' -HeadingType Heading1 -Supress $false

# adding link to paragraph that exists
Add-WordHyperLink -WordDocument $WordDocument -UrlText 'This is Google url' -UrlLink 'https://www.google.com' -Supress $true -Paragraph $Paragraph
# adding link to newly created paragraph
Add-WordHyperLink -WordDocument $WordDocument -UrlText 'This is microsoft url' -UrlLink 'https://www.microsoft.com' -Supress $true

Save-WordDocument $WordDocument -Supress $True -OpenDocument