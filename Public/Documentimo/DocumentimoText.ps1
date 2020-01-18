function DocText {
    [CmdletBinding()]
    [alias('DocumentimoText', 'New-DocumentimoText')]
    param(
        [Parameter(Mandatory = $false, Position = 0)][ScriptBlock] $TextBlock,
        [String[]]$Text,
        [System.Drawing.Color[]]$Color = @(),
        [switch] $LineBreak
    )
    if ($TextBlock) {
        $Text = (Invoke-Command -ScriptBlock $TextBlock)
    }

    [PSCustomObject] @{
        ObjectType = 'Text'
        Text       = $Text
        Color      = $Color
        LineBreak  = $LineBreak
    }
}

