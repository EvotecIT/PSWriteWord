function DocListItem {
    [CmdletBinding()]
    [alias('DocumentimoListItem', 'New-DocumentimoListItem')]
    param(
        [ValidateRange(0, 8)] [int] $Level,
        [string] $Text,
        [nullable[int]] $StartNumber,
        [bool]$TrackChanges = $false,
        [bool]$ContinueNumbering = $false,
        [bool]$Supress = $false
    )

    [PSCustomObject] @{
        ObjectType        = 'ListItem'
        Level             = $Level
        Text              = $Text
        StartNumber       = $StartNumber
        TrackChanges      = $TrackChanges
        ContinueNumbering = $ContinueNumbering
    }
}
