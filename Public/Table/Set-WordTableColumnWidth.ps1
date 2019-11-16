function Set-WordTableColumnWidth {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Document.NET.InsertBeforeOrAfter] $Table,
        [float[]] $Width = @(),
        [nullable[float]] $TotalWidth = $null,
        [bool] $Percentage,
        [bool] $Supress
    )
    if ($null -ne $Table -and $null -ne $Width) {
        if ($Percentage) {
            Write-Verbose "Set-WordTableColumnWidth - Option A - Width: $([string] $Width) - Percentage: $Percentage - TotalWidth: $TotalWidth "
            $Table.SetWidthsPercentage($Width, $TotalWidth)
        } else {
            Write-Verbose "Set-WordTableColumnWidth - Option B - Width: $([string] $Width) - Percentage: $Percentage - TotalWidth: $TotalWidth "
            $Table.SetWidths($Width)
        }
    }
    if ($Supress) { return } else { return $Table }
}