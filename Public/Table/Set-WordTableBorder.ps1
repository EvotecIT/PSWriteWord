function Set-WordTableBorder {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Document.NET.InsertBeforeOrAfter] $Table,
        [nullable[Xceed.Document.NET.TableBorderType]] $TableBorderType,
        $Border,
        [bool] $Supress
    )
    if ($null -ne $Table -and $null -ne $TableBorderType  -and $null -ne $Border) {
        $Table.SetBorder($TableBorderType, $Border)
    }
    if ($Supress) { return } else { $Table }
}