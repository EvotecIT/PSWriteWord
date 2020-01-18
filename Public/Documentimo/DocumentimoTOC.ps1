function DocToc {
    [CmdletBinding()]
    [alias('DocumentimoTOC', 'New-DocumentimoTOC')]
    param(
        [string] $Title,
        [int] $RightTabPos,
        [Xceed.Document.NET.TableOfContentsSwitches] $Switches
    )
    [PSCustomObject] @{
        ObjectType  = 'TOC'
        Title       = $Title
        RightTabPos = $RightTabPos
        Switches    = $Switches
    }
}