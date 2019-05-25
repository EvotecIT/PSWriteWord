Import-Module PSWriteWord.psd1 -Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-AddCustomProperty.docx"

$WordDocument = New-WordDocument -FilePath $FilePath

Add-WordText -WordDocument $WordDocument -Text 'Custom Properties Example' -HeadingType Heading1 -Supress $true

Add-WordCustomProperty -WordDocument $WordDocument -Name 'CompanyName'  -Value 'Evotec' -Supress $true
Add-WordCustomProperty -WordDocument $WordDocument -Name 'CustomEntry'  -Value 'Important Value' -Supress $true

#$PropertyValue = Get-WordCustomProperty -WordDocument $WordDocument -Name 'CompanyName'

$AllCustomProperties = Get-WordCustomProperty -WordDocument $WordDocument

Add-WordText -WordDocument $WordDocument -Text 'Following document has ', $AllCustomProperties.Count, ' custom properties.' -UnderlineStyle none, dash, none -SpacingAfter 5 -Supress $true

foreach ($custom in $AllCustomProperties) {
    Add-WordText -WordDocument $WordDocument -Text 'Custom property: ', $Custom.Name, ' with value: ', $Custom.Value -Bold $false, $true, $false, $true -Supress $true
}

Save-WordDocument $WordDocument -Supress $true

### Start Word with file
Invoke-Item $FilePath