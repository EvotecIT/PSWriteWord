function Add-WordTableRow {
    [CmdletBinding()]
    param (
        [InsertBeforeOrAfter] $Table,
        [int] $Count = 1,
        [nullable[int]] $Index,
        [bool] $Supress = $false
    )
    $List = New-ArrayList
    if ($Table -ne $null) {
        if ($Index -ne $null) {
            for ($i = 0; $i -lt $Count; $i++) {
                #Write-Verbose 'Add-WordTableRow - Adding new row'
                Add-ToArray -List $List -Element $($Table.InsertRow($Index + $i))
            }
        } else {
            for ($i = 0; $i -lt $Count; $i++) {
                #Write-Verbose 'Add-WordTableRow - Adding new row'
                Add-ToArray -List $List -Elemen $($Table.InsertRow())
            }
        }
    }
    if ($Supress) { return } else { return $List }
}