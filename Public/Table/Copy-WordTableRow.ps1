function Copy-WordTableRow {
    [CmdletBinding()]
    param (
        [Xceed.Document.NET.InsertBeforeOrAfter] $Table,
        $Row,
        [nullable[int]] $Index
    )
    if ($Table -ne $null) {
        if ($Index -eq $null) {
            $Table.InsertRow($Row)
        } else {
            $Table.InsertRow($Row, $Index)
        }
    }
}