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
        [alias ("T")] [String[]]$Text,
        [alias ("C")] [System.Drawing.Color[]]$Color = @(),
        [alias ("S")] [double[]] $FontSize = @(),
        [alias ("N")] [string[]] $FontName = @(),
        [alias ("B")] [bool[]] $Bold = @(),
        [alias ("I")] [bool[]] $Italic = @(),
        [alias ("U")] [UnderlineStyle[]] $UnderlineStyle = @(),
        [alias ("SA")] [double[]] $SpacingAfter = @(),
        [alias ("SB")] [double[]] $SpacingBefore = @(),
        [alias ("SP")] [double[]] $Spacing = @(),
        [alias ("H")] [highlight[]] $Highlight = @(),
        [alias ("CA")] [CapsStyle[]] $CapsStyle = @(),
        [alias ("ST")] [StrikeThrough[]] $StrikeThrough = @(),
        [alias ("HT")] [HeadingType[]] $HeadingType = @(),
        [bool[]] $NewLine = @(),
        [switch] $KeepLinesTogether,
        [switch] $KeepWithNextParagraph,
        [bool] $Supress = $true
    )
    if ($Text.Count -eq 0) { return }
    $p = $WordDocument.InsertParagraph()
    for ($i = 0; $i -lt $Text.Length; $i++) {
        if ($NewLine[$i] -ne $null -and $NewLine[$i] -eq $true) {
            if ($i -gt 0) { $p = $WordDocument.InsertParagraph() }
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
        if ($FontName[$i] -ne $null) {
            $p = $p.Font($FontName[$i])
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
        if ($HeadingType[$i] -ne $null) {
            $p = $p.HeadingType($HeadingType[$i])
        }
    }

    if ($Supress) { return } else { return $p }
}