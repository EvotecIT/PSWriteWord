function Add-WordLineChart {
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
        [switch] $NoLegend
    )

    if ($null -eq $ChartSeries) {
        $ChartSeries = Add-WordChartSeries -ChartName $ChartName -Names $Names -Values $Values
    }

    [Xceed.Document.NET.LineChart] $chart = [Xceed.Document.NET.LineChart]::new()
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