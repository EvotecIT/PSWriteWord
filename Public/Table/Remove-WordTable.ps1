function Remove-WordTable {
    [CmdletBinding()]
    param (
        [Xceed.Document.NET.InsertBeforeOrAfter] $Table
    )
    if ($Table -ne $null) {
        $Table.Remove()
    }
}