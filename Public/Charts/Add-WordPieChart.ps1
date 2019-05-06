function Add-WordPieChart {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.Container]$WordDocument,
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [string] $ChartName,
        [string[]] $Names,
        [int[]] $Values,
        [ChartLegendPosition] $ChartLegendPosition = [ChartLegendPosition]::Left,
        [bool] $ChartLegendOverlay = $false,
        [switch] $NoLegend
    )

    $Series = Add-WordChartSeries -ChartName $ChartName -Names $Names -Values $Values

    [Xceed.Words.NET.PieChart] $chart = New-Object -TypeName Xceed.Words.NET.PieChart
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