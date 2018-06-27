function Add-WordTableRow {
    [CmdletBinding()]
    param (
        [Xceed.Words.NET.InsertBeforeOrAfter] $Table,
        [int] $Count = 1,
        [nullable[int]] $Index,
        [bool] $Supress = $true
    )
    $List = New-ArrayList
    if ($Table -ne $null) {
        if ($Index -ne $null) {
            for ($i = 0; $i -lt $Count; $i++) {
                Add-ToArray -List $List -Element $($Table.InsertRow($Index + $i))
            }
        } else {
            for ($i = 0; $i -lt $Count; $i++) {
                Add-ToArray -List $List -Elemen $($Table.InsertRow())
            }
        }
    }
    if ($Supress) { return } else { return $List }
}
function Remove-WordTableRow {
    [CmdletBinding()]
    param (
        [Xceed.Words.NET.InsertBeforeOrAfter] $Table,
        [int] $Count = 1,
        [nullable[int]] $Index
    )
    if ($Table -ne $null) {
        if ($Index -ne $null) {
            for ($i = 0; $i -lt $Count; $i++) {
                $Table.RemoveRow($Index + $i)
            }
        } else {
            for ($i = 0; $i -lt $Count; $i++) {
                $Table.RemoveRow()
            }
        }
    }
}
function Copy-WordTableRow {
    [CmdletBinding()]
    param (
        [Xceed.Words.NET.InsertBeforeOrAfter] $Table,
        $Row,
        [nullable[int]] $Index
    )
    if ($Table -ne $null) {
        if ($Index -eq $null) {
            $Table.InsertRow($Row)
        } else {
            $Table.InsertRow($Row, $Index)
        }
    }
}
function Get-WordTableRow {
    [CmdletBinding()]
    param (
        [Xceed.Words.NET.InsertBeforeOrAfter] $Table,
        [switch] $RowsCount
    )

    if ($Table -ne $null) {
        if ($RowsCount) {
            return $Table.Rows.Count
        }
    }
}