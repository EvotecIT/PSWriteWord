function New-WordTableBorder {
    [CmdletBinding()]
    param (
        [BorderStyle] $BorderStyle,
        [BorderSize] $BorderSize,
        [int] $BorderSpace,
        [System.Drawing.Color] $BorderColor
    )

    $Border = New-Object -TypeName Xceed.Words.NET.Border -ArgumentList $BorderStyle, $BorderSize, $BorderSpace, $BorderColor
    return $Border
}

function Set-WordTableBorder {
    [CmdletBinding()]
    param (
        [Xceed.Words.NET.InsertBeforeOrAfter] $Table,
        [TableBorderType] $TableBorderType,
        $Border
    )
    if ($Table -ne $null -and $TableBorderType -ne $null -and $Border -ne $null) {
        $Table.SetBorder($TableBorderType, $Border)
    }
}
function Set-WordTableDirection {
    [CmdletBinding()]
    param (
        [Xceed.Words.NET.InsertBeforeOrAfter] $Table,
        [Direction] $Direction
    )
    if ($Table -ne $null -and $Direction -ne $null) {
        $Table.SetDirection($Direction)
    }
}

function Set-WordTablePageBreak {
    [CmdletBinding()]
    param (
        [Xceed.Words.NET.InsertBeforeOrAfter] $Table,
        [switch] $AfterTable,
        [switch] $BeforeTable
    )
    if ($Table -ne $null) {
        if ($BeforeTable) {
            $Table.InsertPageBreakBeforeSelf()
        }
        if ($AfterTable) {
            $Table.InsertPageBreakAfterSelf()
        }
    }
}

function Set-WordTable {
    [CmdletBinding()]
    param (
        [Xceed.Words.NET.InsertBeforeOrAfter] $Table,
        [Direction] $Direction,
        [TableBorderType] $TableBorderType,
        $Border
    )


    Set-WordTableDirection -Table $Table -Direction $Direction
    Set-WordTableBorder -Table $Table -TableBorderType $TableBorderType -Border $Border
}