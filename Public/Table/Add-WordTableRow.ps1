function Add-WordTableRow {
    [CmdletBinding()]
    param (
        [InsertBeforeOrAfter] $Table,
        [int] $Count = 1,
        [nullable[int]] $Index,
        [bool] $Supress = $false
    )

    if ($null -ne $Table) {
        $List = @(
            if ($Index -ne $null) {
                for ($i = 0; $i -lt $Count; $i++) {
                    $Table.InsertRow($Index + $i)
                }
            } else {
                for ($i = 0; $i -lt $Count; $i++) {
                    $Table.InsertRow()
                }
            }
        )
    }
    if ($Supress) { return } else { return $List }
}