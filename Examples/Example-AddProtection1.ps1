Import-Module PSWriteWord #-Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-AddProtection1.docx"

$WordDocument = New-WordDocument $FilePath
Add-WordText -WordDocument $WordDocument -Text 'This is text that has font size of 15', ' and this is font size of 10 ', ' while this will be 12.' `
    -FontSize 15, 10 `
    -Color Blue, Red `
    -Bold $true, $false, $true `
    -Italic $true, $true -Supress $True

Add-WordProtection -WordDocument $WordDocument -EditRestrictions readOnly

Save-WordDocument $WordDocument -Supress $True

### Start Word with file
Invoke-Item $FilePath