function Add-WordPieChart {
    [CmdletBinding()]
    param (
        [Xceed.Words.NET.Container]$WordDocument,
        [string] $ChartName,
        [string[]] $Names,
        [int[]] $Values,
        [ChartLegendPosition] $ChartLegendPosition = [ChartLegendPosition]::Left,
        [bool] $ChartLegendOverlay = $false
    )

    $Series = Add-Series -ChartName $ChartName -Names $Names -Values $Values

    [Xceed.Words.NET.PieChart] $chart = New-Object -TypeName Xceed.Words.NET.PieChart
    $chart.AddLegend($ChartLegendPosition, $ChartLegendOverlay)
    $chart.AddSeries($Series)

    $WordDocument.InsertChart($chart)

}

function Add-Series {
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