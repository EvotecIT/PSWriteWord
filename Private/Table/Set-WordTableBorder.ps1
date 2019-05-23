function Set-WordTableBorder {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Table,
        [nullable[Xceed.Words.NET.TableBorderType]] $TableBorderType,
        $Border,
        [bool] $Supress
    )
    if ($null -ne $Table -and $null -ne $TableBorderType  -and $null -ne $Border) {
        $Table.SetBorder($TableBorderType, $Border)
    }
    if ($Supress) { return } else { $Table }
}