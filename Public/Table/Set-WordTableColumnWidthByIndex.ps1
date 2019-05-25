function Set-WordTableColumnWidthByIndex {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][InsertBeforeOrAfter] $Table,
        [nullable[int]] $Index,
        [nullable[double]] $Width
    )
    if ($Table -ne $null -and $Index -ne $null -and $Width -ne $null) {
        $Table.SetColumnWidth($Index, $Width)
    }
}