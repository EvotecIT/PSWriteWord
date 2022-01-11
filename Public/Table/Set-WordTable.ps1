function Set-WordTable {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Document.NET.InsertBeforeOrAfter] $Table,
        [nullable[Xceed.Document.NET.TableBorderType]] $TableBorderType,
        $Border,
        [nullable[Xceed.Document.NET.AutoFit]] $AutoFit,
        [nullable[Xceed.Document.NET.TableDesign]] $Design,
        [nullable[Xceed.Document.NET.Direction]] $Direction,
        [switch] $BreakPageAfterTable,
        [switch] $BreakPageBeforeTable,
        [bool] $Supress
    )
    if ($Table) {
        $Table = $table | Set-WordTableDesign -Design $Design
        $Table = $table | Set-WordTableDirection -Direction $Direction
        $Table = $table | Set-WordTableBorder -TableBorderType $TableBorderType -Border $Border
        $Table = $table | Set-WordTablePageBreak -AfterTable:$BreakPageAfterTable -BeforeTable:$BreakPageBeforeTable
        $Table = $table | Set-WordTableAutoFit -AutoFit $AutoFit
    }
    if ($Supress) { return } Else { return $Table}
}