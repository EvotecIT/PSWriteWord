function Set-WordTableDesign {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Document.NET.InsertBeforeOrAfter] $Table,
        [nullable[Xceed.Document.NET.TableDesign]] $Design
    )
    if ($Table -ne $null -and $Design -ne $null) {
        $Table.Design = $Design
    }
    return $Table
}