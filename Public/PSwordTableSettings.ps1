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
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Table,
        [nullable[TableBorderType]] $TableBorderType,
        $Border,
        [bool] $Supress
    )
    if ($Table -ne $null -and $TableBorderType -ne $null -and $Border -ne $null) {
        $Table.SetBorder($TableBorderType, $Border)
    }
    if ($Supress) { return } else { $Table }
}
function Set-WordTableAutoFit {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Table,
        [nullable[AutoFit]] $AutoFit
    )
    if ($Table -ne $null -and $AutoFit -ne $null) {
        Write-Verbose "Set-WordTabelAutofit - Setting Table Autofit to: $AutoFit"
        $Table.AutoFit = $AutoFit
    }
    return $Table
}

function Set-WordTableDesign {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Table,
        [nullable[TableDesign]] $Design
    )
    if ($Table -ne $null -and $Design -ne $null) {
        $Table.Design = $Design
    }
    return $Table
}

function Set-WordTableDirection {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Table,
        [nullable[Direction]] $Direction
    )
    if ($Table -ne $null -and $Direction -ne $null) {
        $Table.SetDirection($Direction)
    }
    return $Table
}

function Set-WordTablePageBreak {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Table,
        [switch] $AfterTable,
        [switch] $BeforeTable,
        [nullable[bool]] $BreakAcrossPages
    )
    if ($Table -ne $null) {
        if ($BeforeTable) {
            $Table.InsertPageBreakBeforeSelf()
        }
        if ($AfterTable) {
            $Table.InsertPageBreakAfterSelf()
        }
        if ($BreakAcrossPages -ne $null) {
            $Table.BreakAcrossPages = $BreakAcrossPages
        }
    }
    return $Table
}

function Set-WordTable {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Table,
        [nullable[TableBorderType]] $TableBorderType,
        $Border,
        [nullable[AutoFit]] $AutoFit,
        [nullable[TableDesign]] $Design,
        [nullable[Direction]] $Direction,
        [switch] $BreakPageAfterTable,
        [switch] $BreakPageBeforeTable,
        [nullable[bool]] $BreakAcrossPages,
        [bool] $Supress
    )
    if ($Table -ne $null) {
        $Table = $table | Set-WordTableDesign -Design $Design
        $Table = $table | Set-WordTableDirection -Direction $Direction
        $Table = $table | Set-WordTableBorder -TableBorderType $TableBorderType -Border $Border
        $Table = $table | Set-WordTablePageBreak -AfterTable:$BreakPageAfterTable -BeforeTable:$BreakPageBeforeTable -BreakAcrossPages $BreakAcrossPages
        $Table = $table | Set-WordTableAutoFit -AutoFit $AutoFit
    }
    if ($Supress) { return } Else { return $Table}
}