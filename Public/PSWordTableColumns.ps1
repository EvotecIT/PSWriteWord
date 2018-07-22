function Add-WordTableColumn {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Table,
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

function Set-WordTableColumnWidthByIndex {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Table,
        [nullable[int]] $Index,
        [nullable[double]] $Width
    )
    if ($Table -ne $null -and $Index -ne $null -and $Width -ne $null) {
        $Table.SetColumnWidth($Index, $Width)
    }
}

function Set-WordTableColumnWidth {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Table,
        [float[]] $Width = @(),
        [nullable[float]] $TotalWidth = $null,
        [bool] $Percentage,
        [bool] $Supress
    )
    if ($Table -ne $null -and $Width -ne $null) {
        if ($Percentage) {
            Write-Verbose "Set-WordTableColumnWidth - Option A - Width: $([string] $Width) - Percentage: $Percentage - TotalWidth: $TotalWidth "
            $Table.SetWidthsPercentage($Width, $TotalWidth)
        } else {
            Write-Verbose "Set-WordTableColumnWidth - Option B - Width: $([string] $Width) - Percentage: $Percentage - TotalWidth: $TotalWidth "
            $Table.SetWidths($Width)
        }
    }
    if ($Supress) { return } else { return $Table }
}