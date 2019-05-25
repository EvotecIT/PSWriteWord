function Add-WordChartSeries {
    [CmdletBinding()]
    param (
        [string] $ChartName = 'Legend',
        [string[]] $Names,
        [int[]] $Values
    )

    [Array] $rNames = foreach ($Name in $Names) {
        $Name
    }
    [Array] $rValues = foreach ($value in $Values) {
        $value
    }
    [Series] $series = New-Object -TypeName Series -ArgumentList $ChartName
    $Series.Bind($rNames, $rValues)
    return $Series
}