function Set-WordTableDesign {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Table,
        [nullable[Xceed.Words.NET.TableDesign]] $Design
    )
    if ($Table -ne $null -and $Design -ne $null) {
        $Table.Design = $Design
    }
    return $Table
}