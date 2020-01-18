function DocNumbering {
    [CmdletBinding()]
    [alias('DocumentimoNumbering', 'New-DocumentimoNumbering')]
    param(
        [Parameter(Position = 0)][ScriptBlock] $Content,
        [string] $Text,
        [int] $Level = 0,
        [Xceed.Document.NET.ListItemType] $Type = [Xceed.Document.NET.ListItemType]::Numbered,
        [Xceed.Document.NET.HeadingType] $Heading = [Xceed.Document.NET.HeadingType]::Heading1
    )
    [PSCustomObject] @{
        ObjectType = 'TocItem'
        Text       = $Text
        Content    = & $Content
        Level      = $Level
        Type       = $Type
        Heading    = $Heading
    }
}