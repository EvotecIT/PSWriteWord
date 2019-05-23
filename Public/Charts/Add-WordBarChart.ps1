function Add-WordBarChart {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.Container]$WordDocument,
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [string] $ChartName,
        [string[]] $Names,
        [int[]] $Values,
        [Xceed.Words.NET.Series[]] $ChartSeries,
        [Xceed.Words.NET.ChartLegendPosition] $ChartLegendPosition = [Xceed.Words.NET.ChartLegendPosition]::Left,
        [bool] $ChartLegendOverlay = $false,
        [Xceed.Words.NET.BarGrouping] $BarGrouping = [Xceed.Words.NET.BarGrouping]::Standard,
        [Xceed.Words.NET.BarDirection] $BarDirection = [Xceed.Words.NET.BarDirection]::Bar,
        [int] $BarGapWidth = 200,
        [switch] $NoLegend
    )

    if ($null -eq $ChartSeries) {
        $ChartSeries = Add-WordChartSeries -ChartName $ChartName -Names $Names -Values $Values
    }

    [Xceed.Words.NET.BarChart] $chart = New-Object -TypeName Xceed.Words.NET.BarChart
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