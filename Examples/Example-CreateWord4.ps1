Import-Module PSWriteWord #-Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-CreateWord4.docx"

$WordDocument = New-WordDocument $FilePath

Add-WordText -WordDocument $WordDocument -Text 'This is a text' -FontSize 10 -SpacingBefore 50 -UnderlineStyle singleLine -Supress $True
Add-WordText -WordDocument $WordDocument -Text 'This is a text' -FontSize 10 -SpacingBefore 15 -Bold $true -Supress $True
Add-WordText -WordDocument $WordDocument -Text 'This is a text with Heading type 3' -FontSize 15 -HeadingType Heading3 -FontFamily 'Arial' -Italic $true

Add-WordText -WordDocument $WordDocument -Text 'This is a text', ' that will show ', 'how Add-WordText works ', 'without', ' continue formatting feature.' -FontFamily Tahoma -FontSize 10 -Color Blue
Add-WordText -WordDocument $WordDocument -Text 'This is a text', ' that will show ', 'how Add-WordText works ', 'with...', ' continue formatting feature.' -FontFamily Tahoma -FontSize 10 -Color Blue -ContinueFormatting #-Verbose

Save-WordDocument $WordDocument -Language 'en-US'

### Start Word with file
Invoke-Item $FilePath