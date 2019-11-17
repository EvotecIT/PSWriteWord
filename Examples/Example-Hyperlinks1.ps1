Import-Module .\PSWriteWord.psd1 -Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-HyperLinks1.docx"

$WordDocument = New-WordDocument -FilePath $FilePath

Add-WordText -WordDocument $WordDocument -Text 'This is my first title' -HeadingType Heading1 -Supress $True
Add-WordHyperLink -WordDocument $WordDocument -UrlText 'This is my url' -UrlLink 'https://evotec.xyz' -Color Blue -Alignment center -UnderlineColor Red -UnderlineStyle dotted -Italic $true -Supress $true
Add-WordHyperLink -WordDocument $WordDocument -UrlText 'This is my url' -UrlLink 'https://evotec.xyz' -Color Brown -Bold $true -Italic $true -CapsStyle caps -Supress $true
Save-WordDocument $WordDocument -Supress $True -OpenDocument