function Add-WordTableRow {
    [CmdletBinding()]
    param (
        [Xceed.Words.NET.InsertBeforeOrAfter] $Table,
        [int] $Count = 1,
        [nullable[int]] $Index
    )
    if ($Table -ne $null) {
        if ($Index -ne $null) {
            for ($i = 0; $i -lt $Count; $i++) {
                $Table.InsertRow($Index + $i)
            }
        } else {
            for ($i = 0; $i -lt $Count; $i++) {
                $Table.InsertRow()
            }
        }
    }
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
