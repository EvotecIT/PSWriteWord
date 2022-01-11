function Set-WordTablePageBreak {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Document.NET.InsertBeforeOrAfter] $Table,
        [switch] $AfterTable,
        [switch] $BeforeTable
    )
    if ($Table) {
        if ($BeforeTable) {
            $Table.InsertPageBreakBeforeSelf()
        }
        if ($AfterTable) {
            $Table.InsertPageBreakAfterSelf()
        }
    }
    return $Table
}