function Add-WordTableColumn {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Document.NET.InsertBeforeOrAfter] $Table,
        [int] $Count = 1,
        [int] $Index,
        [ValidateSet('Left', 'Right')] $Direction = 'Left'
    )
    if ($Direction -eq 'Left') { $ColumnSide = $false } else { $ColumnSide = $true }
    if ($null -ne $Table) {
        #  if ($Index -ne $null) {
        for ($i = 0; $i -lt $Count; $i++) {
            $Table.InsertColumn($Index + $i, $ColumnSide)
        }
        #   } else {
        #   for ($i = 0; $i -lt $Count; $i++) {
        #      $Table.InsertColumn()
        #   }
        #     }
    }
}