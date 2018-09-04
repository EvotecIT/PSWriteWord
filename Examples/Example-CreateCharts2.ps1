Import-Module PSWriteWord #-Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-CreateCharts2.docx"

$WordDocument = New-WordDocument $FilePath
Add-WordText -WordDocument $WordDocument -Text 'Line Chart Example #1' `
    -FontSize 15 `
    -Color Blue `
    -Bold $true -HeadingType Heading1 -Supress $True

Add-WordLineChart -WordDocument $WordDocument -ChartName 'My finances' -Names 'Today', 'Yesterday', 'Two days ago' -Values 1050.50, 2000, 20000 -ChartLegendPosition Bottom -ChartLegendOverlay $false


Add-WordText -WordDocument $WordDocument -Text 'Line Chart Example #2' `
    -FontSize 15 `
    -Color Blue `
    -Bold $true -HeadingType Heading1 -Supress $True


$Series1 = Add-WordChartSeries -ChartName 'One'  -Names 'Today', 'Yesterday', 'Two days ago' -Values 1050.50, 2000, 20000
$Series2 = Add-WordChartSeries -ChartName 'Two'  -Names 'Today', 'Yesterday', 'Two days ago' -Values 3000, 2000, 1000

Add-WordLineChart -WordDocument $WordDocument -ChartName 'My finances' -ChartLegendPosition Bottom -ChartLegendOverlay $false -ChartSeries $Series1, $Series2

Add-WordText -WordDocument $WordDocument -Text 'Line Chart Example #3 - No legend' `
    -FontSize 15 `
    -Color Blue `
    -Bold $true -HeadingType Heading1 -Supress $True

Add-WordText -WordDocument $WordDocument -Text "Keep in mind that you need to define new series for each new chart. Otherwise errors will occur..." -Supress $True

$Series3 = Add-WordChartSeries -ChartName 'One'  -Names 'Today', 'Yesterday', 'Two days ago' -Values 1050.50, 2000, 20000
$Series4 = Add-WordChartSeries -ChartName 'Two'  -Names 'Today', 'Yesterday', 'Two days ago' -Values 3000, 2000, 1000

Add-WordLineChart -WordDocument $WordDocument -ChartName 'My finances1' -ChartSeries $Series3, $Series4 -NoLegend


Save-WordDocument $WordDocument -Supress $True

### Start Word with file
Invoke-Item $FilePath