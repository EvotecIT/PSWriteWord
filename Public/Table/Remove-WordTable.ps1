function Remove-WordTable {
    [CmdletBinding()]
    param (
        [InsertBeforeOrAfter] $Table
    )
    if ($Table -ne $null) {
        $Table.Remove()
    }
}