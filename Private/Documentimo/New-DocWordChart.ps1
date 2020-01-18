function New-DocWordChart {
    [CmdletBinding()]
    param(
        [Xceed.Document.NET.Container] $WordDocument,
        [PSCustomObject] $Parameters
    )

    [Array] $DataTable = $Parameters.DataTable
    [string] $Key = $Parameters.Key
    [string] $Value = $Parameters.Value
    [Xceed.Document.NET.ChartLegendPosition] $LegendPosition = $Parameters.LegendPosition
    [bool] $LegendOverlay = $Parameters.LegendOverlay
    [string] $Title = $Parameters.Title

    if ($DataTable.Count -gt 0) {
        # If chart had no values or keys it would create an empty chart and prevent saving of document in Word
        if ($DataTable[0] -is [System.Collections.IDictionary]) {
            $TemporaryNames = $DataTable.Keys
            $Names = foreach ($Name in $TemporaryNames) {
                "$Name ($($DataTable.$Name))"
            }
            $Values = $DataTable.Values
        } else {
            if (($Key -ne '') -and ($Value -ne '')) {
                $Names = for ($i = 0; $i -lt $DataTable.Count; $i++) {
                    "$($DataTable[$i].$Key) ($($DataTable[$i].$Value))"
                }
                $Values = $DataTable.$Value
            } else {
                return
            }
        }
        if (($Values.Count -eq $Names.Count) -and ($Values.Count -gt 0)) {
            Add-WordParagraph -WordDocument $WordDocument -Supress $True

            Add-WordPieChart -WordDocument $WordDocument -ChartName $Title -Names $Names -Values $Values -ChartLegendPosition $LegendPosition -ChartLegendOverlay $LegendOverlay
        } else {
            Write-Warning "DocumentimoChart - Names and Values count doesn't match or equals 0"
        }
    }
}