function Add-WordTableColumn {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][InsertBeforeOrAfter] $Table,
        [int] $Count = 1,
        [nullable[int]] $Index,
        [ValidateSet('Left', 'Right')] $Direction = 'Left'
    )
    if ($Direction -eq 'Left') { $ColumnSide = $false} else { $ColumnSide = $true}

    if ($Table -ne $null) {
        if ($Index -ne $null) {
            for ($i = 0; $i -lt $Count; $i++) {
                $Table.InsertColumn($Index + $i, $ColumnSide)
            }
        } else {
            for ($i = 0; $i -lt $Count; $i++) {
                $Table.InsertColumn()
            }
        }
    }
}