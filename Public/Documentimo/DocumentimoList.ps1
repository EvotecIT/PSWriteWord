function DocList {
    [CmdletBinding()]
    [alias('DocumentimoList', 'New-DocumentimoList')]
    param(
        [ScriptBlock] $ListItems,
        [alias('ListType')][Xceed.Document.NET.ListItemType] $Type = [Xceed.Document.NET.ListItemType]::Bulleted
    )

    [PSCustomObject] @{
        ObjectType = 'List'
        ListItems  = Invoke-Command -ScriptBlock $ListItems
        Type       = $Type
    }
}