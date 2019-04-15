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