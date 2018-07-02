Import-Module PSWriteWord #-Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-AddProtection2.docx"

$WordDocument = New-WordDocument $FilePath
Add-WordText -WordDocument $WordDocument -Text 'This is text that has font size of 15', ' and this is font size of 10 ', ' while this will be 12.' `
    -FontSize 15, 10 `
    -Color Blue, Red `
    -Bold $true, $false, $true `
    -Italic $true, $true

Add-WordProtection -WordDocument $WordDocument -EditRestrictions readOnly -Password '12345678'

Save-WordDocument $WordDocument

### Start Word with file
Invoke-Item $FilePath