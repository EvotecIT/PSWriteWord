# Paragraph InsertParagraph()
# Paragraph InsertParagraph( int index, string text, bool trackChanges )
# Paragraph InsertParagraph( Paragraph p )
# Paragraph InsertParagraph( int index, Paragraph p )
# Paragraph InsertParagraph( int index, string text, bool trackChanges, Formatting formatting )
# Paragraph InsertParagraph( string text )
# Paragraph InsertParagraph( string text, bool trackChanges )
# Paragraph InsertParagraph( string text, bool trackChanges, Formatting formatting )

function Add-WordText {
    [CmdletBinding()]
    param (
        [Xceed.Words.NET.Container]$WordDocument,
        [Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [alias ("T")] [String[]]$Text,
        [alias ("C")] [System.Drawing.Color[]]$Color = @(),
        [alias ("S")] [double[]] $FontSize = @(),
        [alias ("FontName")] [string[]] $FontFamily = @(),
        [alias ("B")] [bool[]] $Bold = @(),
        [alias ("I")] [bool[]] $Italic = @(),
        [alias ("U")] [UnderlineStyle[]] $UnderlineStyle = @(),
        [alias ('UC')] [System.Drawing.Color[]]$UnderlineColor = @(),
        [alias ("SA")] [double[]] $SpacingAfter = @(),
        [alias ("SB")] [double[]] $SpacingBefore = @(),
        [alias ("SP")] [double[]] $Spacing = @(),
        [alias ("H")] [highlight[]] $Highlight = @(),
        [alias ("CA")] [CapsStyle[]] $CapsStyle = @(),
        [alias ("ST")] [StrikeThrough[]] $StrikeThrough = @(),
        [alias ("HT")] [HeadingType[]] $HeadingType = @(),
        $PercentageScale = @(), # "Value must be one of the following: 200, 150, 100, 90, 80, 66, 50 or 33"
        $Misc = @(),
        [string[]] $Language = @(),
        $Kerning = @(), # "Value must be one of the following: 8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48 or 72"
        $Hidden = @(),
        $Position = @(), #  "Value must be in the range -1585 - 1585"
        [bool[]] $NewLine = @(),
        [switch] $KeepLinesTogether,
        [switch] $KeepWithNextParagraph,
        [single[]] $IndentationFirstLine = @(),
        [single[]] $IndentationHanging = @(),
        [Alignment[]] $Alignment = @(),
        [bool] $Supress = $true
    )
    if ($Text.Count -eq 0) { return }

    #if ($Paragraph -eq $null) {

    #}
    $p = $WordDocument.InsertParagraph()
    if ($Paragraph -ne $null) {
        $p = $Paragraph.InsertParagraphAfterSelf($p)
    }

    for ($i = 0; $i -lt $Text.Length; $i++) {
        if ($NewLine[$i] -ne $null -and $NewLine[$i] -eq $true) {
            if ($i -gt 0) {
                if ($Paragraph -ne $null) {
                    $p = $p.InsertParagraphAfterSelf()
                } else {
                    $p = $WordDocument.InsertParagraph()
                }
            }
            $p = $p.Append($Text[$i])
        } else {
            $p = $p.Append($Text[$i])
        }
        if ($Color[$i] -ne $null) {
            $p = $p.Color($Color[$i])
        }
        if ($FontSize[$i] -ne $null) {
            $p = $p.FontSize($FontSize[$i])
        }
        if ($FontFamily[$i] -ne $null) {
            $p = $p.Font($FontFamily[$i])
        }
        if ($Bold[$i] -ne $null -and $Bold[$i] -eq $true) {
            $p = $p.Bold()
        }
        if ($Italic[$i] -ne $null -and $Italic[$i] -eq $true) {
            $p = $p.Italic()
        }
        if ($UnderlineStyle[$i] -ne $null) {
            $p = $p.UnderlineStyle($UnderlineStyle[$i])
        }
        if ($UnderlineColor[$i] -ne $null) {
            $p = $p.UnderlineColor($UnderlineColor[$i])
        }
        if ($SpacingAfter[$i] -ne $null) {
            $p = $p.SpacingAfter($SpacingAfter[$i])
        }
        if ($SpacingBefore[$i] -ne $null) {
            $p = $p.SpacingBefore($SpacingBefore[$i])
        }
        if ($SpacingBefore[$i] -ne $null) {
            $p = $p.SpacingBefore($SpacingBefore[$i])
        }
        if ($Spacing[$i] -ne $null) {
            $p = $p.Spacing($Spacing[$i])
        }
        if ($Highlight[$i] -ne $null) {
            $p = $p.Highlight($Highlight[$i])
        }
        if ($CapsStyle[$i] -ne $null) {
            $p = $p.CapsStyle($CapsStyle[$i])
        }
        if ($StrikeThrough[$i] -ne $null) {
            $p = $p.StrikeThrough($StrikeThrough[$i])
        }
        if ($PercentageScale[$i] -ne $null) {
            $p = $p.PercentageScale($PercentageScale[$i])
        }
        if ($Language[$i] -ne $null) {
            Write-Verbose "Add-WriteText - Setting language $($Language[$i])"
            $Culture = [System.Globalization.CultureInfo]::GetCultureInfo($Language[$i])
            $p = $p.Culture($Culture)
        }
        if ($Kerning[$i] -ne $null) {
            $p = $p.Kerning($Kerning[$i])
        }
        if ($PercentageScale[$i] -ne $null) {
            $p = $p.PercentageScale($PercentageScale[$i])
        }
        if ($Misc[$i] -ne $null) {
            $p = $p.Misc($Misc[$i])
        }
        if ($Position[$i] -ne $null) {
            $p = $p.Position($Position[$i])
        }
        if ($HeadingType[$i] -ne $null) {
            $p.StyleName = $HeadingType[$i]
        }
        if ($Alignment[$i] -ne $null) {
            $p.Alignment = $Alignment[$i]
        }
        if ($IndentationFirstLine[$i] -ne $null) {
            $p.IndentationFirstLine = $IndentationFirstLine[$i]
        }
        if ($IndentationHanging[$i] -ne $null) {
            $p.IndentationHanging = $IndentationHanging[$i]
        }
    }

    if ($Supress) { return } else { return $p }
}