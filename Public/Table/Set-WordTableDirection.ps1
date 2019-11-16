function Set-WordTableDirection {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Document.NET.InsertBeforeOrAfter] $Table,
        [nullable[Xceed.Document.NET.Direction]] $Direction
    )
    if ($Table -ne $null -and $Direction -ne $null) {
        $Table.SetDirection($Direction)
    }
    return $Table
}