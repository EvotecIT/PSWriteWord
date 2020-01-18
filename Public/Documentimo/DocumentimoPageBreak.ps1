function DocPageBreak {
    [CmdletBinding()]
    [alias('DocumentimoPageBreak', 'New-DocumentimoPageBreak')]
    param(
        [int] $Count = 1
    )

    [PSCustomObject] @{
        ObjectType = 'PageBreak'
        Count      = $Count
    }
}

