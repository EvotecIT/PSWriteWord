function Set-WordTableDirection {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][InsertBeforeOrAfter] $Table,
        [nullable[Direction]] $Direction
    )
    if ($Table -ne $null -and $Direction -ne $null) {
        $Table.SetDirection($Direction)
    }
    return $Table
}