Import-Module PSWriteWord #-Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-CreateCharts1.docx"

$WordDocument = New-WordDocument $FilePath
Add-WordText -WordDocument $WordDocument -Text 'Pie Chart Example #1' `
    -FontSize 15 `
    -Color Blue `
    -Bold $true -HeadingType Heading1

Add-WordPieChart -WordDocument $WordDocument -ChartName 'My finances' -Names 'Today', 'Yesterday', 'Two days ago' -Values 1050.50, 2000, 20000 -ChartLegendPosition Bottom -ChartLegendOverlay $false

Add-WordText -WordDocument $WordDocument -Text 'Pie Chart Example #2' `
    -FontSize 15 `
    -Color Blue `
    -Bold $true -HeadingType Heading1

Add-WordPieChart -WordDocument $WordDocument -ChartName 'My finances' -Names 'Today', 'Yesterday' -Values  2000, 20000 -ChartLegendPosition Left -ChartLegendOverlay $true

Save-WordDocument $WordDocument -Language 'en-US'

### Start Word with file
Invoke-Item $FilePath