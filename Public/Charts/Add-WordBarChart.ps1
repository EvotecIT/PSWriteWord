function Add-WordBarChart {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Document.NET.Container]$WordDocument,
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Document.NET.InsertBeforeOrAfter] $Paragraph,
        [string] $ChartName,
        [string[]] $Names,
        [int[]] $Values,
        [Xceed.Document.NET.Series[]] $ChartSeries,
        [Xceed.Document.NET.ChartLegendPosition] $ChartLegendPosition = [Xceed.Document.NET.ChartLegendPosition]::Left,
        [bool] $ChartLegendOverlay = $false,
        [Xceed.Document.NET.BarGrouping] $BarGrouping = [Xceed.Document.NET.BarGrouping]::Standard,
        [Xceed.Document.NET.BarDirection] $BarDirection = [Xceed.Document.NET.BarDirection]::Bar,
        [int] $BarGapWidth = 200,
        [switch] $NoLegend
    )

    if ($null -eq $ChartSeries) {
        $ChartSeries = Add-WordChartSeries -ChartName $ChartName -Names $Names -Values $Values
    }

    [Xceed.Document.NET.BarChart] $chart = [Xceed.Document.NET.BarChart]::new()
    $chart.BarDirection = $BarDirection
    $chart.BarGrouping = $BarGrouping
    $chart.GapWidth = $BarGapWidth
    if (-not $NoLegend) {
        $chart.AddLegend($ChartLegendPosition, $ChartLegendOverlay)
    }
    foreach ($series in $ChartSeries) {
        $chart.AddSeries($Series)
    }
    if ($Paragraph -eq $null) {
        $WordDocument.InsertChart($chart)
    } else {
        $WordDocument.InsertChartAfterParagraph($chart, $paragraph)
    }
}