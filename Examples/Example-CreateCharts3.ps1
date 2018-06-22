Import-Module PSWriteWord #-Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-CreateCharts3.docx"

$WordDocument = New-WordDocument $FilePath
Add-WordText -WordDocument $WordDocument -Text 'Bar Chart Example #1' `
    -FontSize 15 `
    -Color Blue `
    -Bold $true -HeadingType Heading1

Add-WordBarChart -WordDocument $WordDocument -ChartName 'My finances' -Names 'Today', 'Yesterday', 'Two days ago' -Values 1050.50, 2000, 20000 -ChartLegendPosition Bottom -ChartLegendOverlay $false

Add-WordText -WordDocument $WordDocument -Text 'Bar Chart Example #2' `
    -FontSize 15 `
    -Color Blue `
    -Bold $true -HeadingType Heading1

$Series1 = Add-WordChartSeries -ChartName 'One'  -Names 'Today', 'Yesterday', 'Two days ago' -Values 1050.50, 2000, 20000
$Series2 = Add-WordChartSeries -ChartName 'Two'  -Names 'Today', 'Yesterday', 'Two days ago' -Values 3000, 2000, 1000

Add-WordBarChart -WordDocument $WordDocument -ChartName 'My finances'-ChartLegendPosition Bottom -ChartLegendOverlay $false -ChartSeries $Series1, $Series2

Add-WordText -WordDocument $WordDocument -Text 'Bar Chart Example #3' `
    -FontSize 15 `
    -Color Blue `
    -Bold $true -HeadingType Heading1


$Series3 = Add-WordChartSeries -ChartName 'One'  -Names 'Today', 'Yesterday', 'Two days ago' -Values 1050.50, 2000, 20000
$Series4 = Add-WordChartSeries -ChartName 'Two'  -Names 'Today', 'Yesterday', 'Two days ago' -Values 3000, 2000, 1000


Add-WordBarChart -WordDocument $WordDocument -ChartName 'My finances'-ChartLegendPosition Bottom -ChartLegendOverlay $false -ChartSeries $Series3, $Series4 -BarGrouping Stacked

Add-WordText -WordDocument $WordDocument -Text 'Bar Chart Example #4' `
    -FontSize 15 `
    -Color Blue `
    -Bold $true -HeadingType Heading1


$Series5 = Add-WordChartSeries -ChartName 'One'  -Names 'Today', 'Yesterday', 'Two days ago' -Values 1050.50, 2000, 20000
$Series6 = Add-WordChartSeries -ChartName 'Two'  -Names 'Today', 'Yesterday', 'Two days ago' -Values 3000, 2000, 1000

Add-WordBarChart -WordDocument $WordDocument -ChartName 'My finances'-ChartLegendPosition Bottom -ChartLegendOverlay $false -ChartSeries $Series5, $Series6 -BarDirection Column

Save-WordDocument $WordDocument