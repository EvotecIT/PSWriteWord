function Add-WordPieChart {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.Container]$WordDocument,
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [string] $ChartName,
        [string[]] $Names,
        [int[]] $Values,
        [ChartLegendPosition] $ChartLegendPosition = [ChartLegendPosition]::Left,
        [bool] $ChartLegendOverlay = $false
    )

    $Series = Add-WordChartSeries -ChartName $ChartName -Names $Names -Values $Values

    [Xceed.Words.NET.PieChart] $chart = New-Object -TypeName Xceed.Words.NET.PieChart
    $chart.AddLegend($ChartLegendPosition, $ChartLegendOverlay)
    $chart.AddSeries($Series)

    if ($Paragraph -eq $null) {
        $WordDocument.InsertChart($chart)
    } else {
        $WordDocument.InsertChartAfterParagraph($chart, $paragraph)
    }

}

function Add-WordLineChart {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.Container]$WordDocument,
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [string] $ChartName,
        [string[]] $Names,
        [int[]] $Values,
        [Xceed.Words.NET.Series[]] $ChartSeries,
        [ChartLegendPosition] $ChartLegendPosition = [ChartLegendPosition]::Left,
        [bool] $ChartLegendOverlay = $false
    )

    if ($ChartSeries -eq $null) {
        $ChartSeries = Add-WordChartSeries -ChartName $ChartName -Names $Names -Values $Values
    }

    [Xceed.Words.NET.LineChart] $chart = New-Object -TypeName Xceed.Words.NET.LineChart
    $chart.AddLegend($ChartLegendPosition, $ChartLegendOverlay)
    foreach ($series in $ChartSeries) {
        $chart.AddSeries($Series)
    }
    if ($Paragraph -eq $null) {
        $WordDocument.InsertChart($chart)
    } else {
        $WordDocument.InsertChartAfterParagraph($chart, $paragraph)
    }
}

function Add-WordBarChart {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.Container]$WordDocument,
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [string] $ChartName,
        [string[]] $Names,
        [int[]] $Values,
        [Xceed.Words.NET.Series[]] $ChartSeries,
        [ChartLegendPosition] $ChartLegendPosition = [ChartLegendPosition]::Left,
        [bool] $ChartLegendOverlay = $false,
        [BarGrouping] $BarGrouping = [BarGrouping]::Standard,
        [BarDirection] $BarDirection = [BarDirection]::Bar,
        [int] $BarGapWidth = 200
    )

    if ($ChartSeries -eq $null) {
        $ChartSeries = Add-WordChartSeries -ChartName $ChartName -Names $Names -Values $Values
    }

    [Xceed.Words.NET.BarChart] $chart = New-Object -TypeName Xceed.Words.NET.BarChart
    $chart.BarDirection = $BarDirection
    $chart.BarGrouping = $BarGrouping
    $chart.GapWidth = $BarGapWidth
    $chart.AddLegend($ChartLegendPosition, $ChartLegendOverlay)
    foreach ($series in $ChartSeries) {
        $chart.AddSeries($Series)
    }
    if ($Paragraph -eq $null) {
        $WordDocument.InsertChart($chart)
    } else {
        $WordDocument.InsertChartAfterParagraph($chart, $paragraph)
    }
}
function Add-WordChartSeries {
    param (
        [string] $ChartName = 'Legend',
        [string[]] $Names,
        [int[]] $Values
    )

    $rNames = New-Object "System.Collections.Generic.List[string]"
    $rValues = New-Object "System.Collections.Generic.List[int]"
    foreach ($name in $names) {
        $rNames.Add($name)
    }
    foreach ($value in $values) {
        $rValues.Add($value)

    }
    [Xceed.Words.NET.Series] $series = New-Object -TypeName Xceed.Words.NET.Series -ArgumentList $ChartName
    $Series.Bind($rNames, $rValues)
    return $Series
}