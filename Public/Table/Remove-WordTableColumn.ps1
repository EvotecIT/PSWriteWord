function Remove-WordTableColumn {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][InsertBeforeOrAfter] $Table,
        [int] $Count = 1,
        [nullable[int]] $Index
    )
    if ($Table) {
        if ($Index) {
            for ($i = 0; $i -lt $Count; $i++) {
                $Table.RemoveColumn($Index + $i)
            }
        } else {
            for ($i = 0; $i -lt $Count; $i++) {
                $Table.RemoveColumn()
            }
        }
    }
}
