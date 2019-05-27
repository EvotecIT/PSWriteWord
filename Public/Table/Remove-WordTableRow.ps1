function Remove-WordTableRow {
    [CmdletBinding()]
    param (
        [InsertBeforeOrAfter] $Table,
        [int] $Count = 1,
        [nullable[int]] $Index,
        [bool] $Supress
    )
    if ($Table) {
        if ($Index) {
            for ($i = 0; $i -lt $Count; $i++) {
                $Table.RemoveRow($Index + $i)
            }
        } else {
            for ($i = 0; $i -lt $Count; $i++) {
                $Table.RemoveRow()
            }
        }
    }
    if ($Supress) { return } else { return $Table}
}