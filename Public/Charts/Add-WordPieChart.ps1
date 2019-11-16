function Add-WordPieChart {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Document.NET.Container]$WordDocument,
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Document.NET.InsertBeforeOrAfter] $Paragraph,
        [string] $ChartName,
        [string[]] $Names,
        [int[]] $Values,
        [Xceed.Document.NET.ChartLegendPosition] $ChartLegendPosition = [Xceed.Document.NET.ChartLegendPosition]::Left,
        [bool] $ChartLegendOverlay = $false,
        [switch] $NoLegend
    )

    $Series = Add-WordChartSeries -ChartName $ChartName -Names $Names -Values $Values

    [Xceed.Document.NET.PieChart] $chart = [Xceed.Document.NET.PieChart]::new()
    if (-not $NoLegend) {
        $chart.AddLegend($ChartLegendPosition, $ChartLegendOverlay)
    }
    $chart.AddSeries($Series)

    if ($null -eq $Paragraph) {
        $WordDocument.InsertChart($chart)
    } else {
        $WordDocument.InsertChartAfterParagraph($chart, $paragraph)
    }
}