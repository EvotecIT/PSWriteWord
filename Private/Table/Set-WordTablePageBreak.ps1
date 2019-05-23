function Set-WordTablePageBreak {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Table,
        [switch] $AfterTable,
        [switch] $BeforeTable,
        [nullable[bool]] $BreakAcrossPages
    )
    if ($Table) {
        if ($BeforeTable) {
            $Table.InsertPageBreakBeforeSelf()
        }
        if ($AfterTable) {
            $Table.InsertPageBreakAfterSelf()
        }
        if ($BreakAcrossPages -ne $null) {
            $Table.BreakAcrossPages = $BreakAcrossPages
        }
    }
    return $Table
}