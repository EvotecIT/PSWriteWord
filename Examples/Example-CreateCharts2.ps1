Import-Module PSWriteWord -Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-CreateCharts2.docx"

$WordDocument = New-WordDocument $FilePath
Add-WordText -WordDocument $WordDocument -Text 'Line Chart Example #1' `
    -FontSize 15 `
    -Color Blue `
    -Bold $true -HeadingType Heading1

Add-WordLineChart -WordDocument $WordDocument -ChartName 'My finances' -Names 'Today', 'Yesterday', 'Two days ago' -Values 1050.50, 2000, 20000 -ChartLegendPosition Bottom -ChartLegendOverlay $false


Add-WordText -WordDocument $WordDocument -Text 'Line Chart Example #2' `
    -FontSize 15 `
    -Color Blue `
    -Bold $true -HeadingType Heading1


$Series1 = Add-WordChartSeries -ChartName 'One'  -Names 'Today', 'Yesterday', 'Two days ago' -Values 1050.50, 2000, 20000
$Series2 = Add-WordChartSeries -ChartName 'Two'  -Names 'Today', 'Yesterday', 'Two days ago' -Values 3000, 2000, 1000

Add-WordLineChart -WordDocument $WordDocument -ChartName 'My finances'-ChartLegendPosition Bottom -ChartLegendOverlay $false -ChartSeries $Series1, $Series2


Save-WordDocument $WordDocument