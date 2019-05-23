function Set-WordTable {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Table,
        [nullable[Xceed.Words.NET.TableBorderType]] $TableBorderType,
        $Border,
        [nullable[Xceed.Words.NET.AutoFit]] $AutoFit,
        [nullable[Xceed.Words.NET.TableDesign]] $Design,
        [nullable[Xceed.Words.NET.Direction]] $Direction,
        [switch] $BreakPageAfterTable,
        [switch] $BreakPageBeforeTable,
        [nullable[bool]] $BreakAcrossPages,
        [bool] $Supress
    )
    if ($Table -ne $null) {
        $Table = $table | Set-WordTableDesign -Design $Design
        $Table = $table | Set-WordTableDirection -Direction $Direction
        $Table = $table | Set-WordTableBorder -TableBorderType $TableBorderType -Border $Border
        $Table = $table | Set-WordTablePageBreak -AfterTable:$BreakPageAfterTable -BeforeTable:$BreakPageBeforeTable -BreakAcrossPages $BreakAcrossPages
        $Table = $table | Set-WordTableAutoFit -AutoFit $AutoFit
    }
    if ($Supress) { return } Else { return $Table}
}