function Set-WordTableDesign {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][InsertBeforeOrAfter] $Table,
        [nullable[TableDesign]] $Design
    )
    if ($Table -ne $null -and $Design -ne $null) {
        $Table.Design = $Design
    }
    return $Table
}