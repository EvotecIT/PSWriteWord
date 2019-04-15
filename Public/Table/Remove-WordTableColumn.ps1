function Remove-WordTableColumn {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Table,
        [int] $Count = 1,
        [nullable[int]] $Index
    )
    if ($Table -ne $null) {
        if ($Index -ne $null) {
            for ($i = 0; $i -lt $Count; $i++) {
                $Table.RemoveColumn($Index + $i)
            }
        } else {
            for ($i = 0; $i -lt $Count; $i++) {
                $Table.RemoveColumn()
            }
        }
    }
}
