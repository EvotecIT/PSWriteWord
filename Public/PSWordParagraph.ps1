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
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.Container]$WordDocument,
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
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
        [int[]] $PercentageScale = @(), # "Value must be one of the following: 200, 150, 100, 90, 80, 66, 50 or 33"
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
        [Direction[]] $Direction = @(),
        [bool] $Supress = $true
    )
    if ($Text.Count -eq 0) { return }

    if ($Paragraph -ne $null) {
        $Paragraph = $Paragraph.InsertParagraphAfterSelf($Paragraph)
    } else {
        if ($WordDocument -ne $null) {
            $Paragraph = $WordDocument.InsertParagraph()
        } else {
            throw 'Both Paragraph and WordDocument are null'
        }
    }
    for ($i = 0; $i -lt $Text.Length; $i++) {
        if ($NewLine[$i] -ne $null -and $NewLine[$i] -eq $true) {
            if ($i -gt 0) {
                if ($Paragraph -ne $null) {
                    $Paragraph = $Paragraph.InsertParagraphAfterSelf()
                } else {
                    $Paragraph = $WordDocument.InsertParagraph()
                }
            }
            $Paragraph = $Paragraph.Append($Text[$i])
        } else {
            $Paragraph = $Paragraph.Append($Text[$i])
        }
        $Paragraph = $Paragraph | Set-WordTextColor -Color $Color[$i] -Supress $false
        $Paragraph = $Paragraph | Set-WordTextFontSize -FontSize $FontSize[$i] -Supress $false
        $Paragraph = $Paragraph | Set-WordTextFontFamily -FontFamily $FontFamily[$i] -Supress $false
        $Paragraph = $Paragraph | Set-WordTextBold -Bold $Bold[$i] -Supress $false
        $Paragraph = $Paragraph | Set-WordTextItalic -Italic $Italic[$i] -Supress $false
        $Paragraph = $Paragraph | Set-WordTextUnderlineColor -UnderlineColor $UnderlineColor[$i] -Supress $false
        $Paragraph = $Paragraph | Set-WordTextUnderlineStyle -UnderlineStyle $UnderlineStyle[$i] -Supress $false
        $Paragraph = $Paragraph | Set-WordTextSpacingAfter -SpacingAfter $SpacingAfter[$i] -Supress $false
        $Paragraph = $Paragraph | Set-WordTextSpacingBefore -SpacingBefore $SpacingBefore[$i] -Supress $false
        $Paragraph = $Paragraph | Set-WordTextSpacing -Spacing $Spacing[$i] -Supress $false
        $Paragraph = $Paragraph | Set-WordTextHighlight -Highlight $Highlight[$i] -Supress $false
        $Paragraph = $Paragraph | Set-WordTextCapsStyle -CapsStyle $CapsStyle[$i] -Supress $false
        $Paragraph = $Paragraph | Set-WordTextStrikeThrough -StrikeThrough $StrikeThrough[$i] -Supress $false
        $Paragraph = $Paragraph | Set-WordTextPercentageScale -PercentageScale $PercentageScale[$i] -Supress $false
        $Paragraph = $Paragraph | Set-WordTextSpacing -Spacing $Spacing[$i] -Supress $false
        $Paragraph = $Paragraph | Set-WordTextLanguage -Language $Language[$i] -Supress $false

        if ($Kerning[$i] -ne $null) {
            $Paragraph = $Paragraph.Kerning($Kerning[$i])
        }
        if ($Misc[$i] -ne $null) {
            $Paragraph = $Paragraph.Misc($Misc[$i])
        }
        if ($Position[$i] -ne $null) {
            $Paragraph = $Paragraph.Position($Position[$i])
        }
        if ($HeadingType[$i] -ne $null) {
            $Paragraph.StyleName = $HeadingType[$i]
        }
        if ($Alignment[$i] -ne $null) {
            $Paragraph.Alignment = $Alignment[$i]
        }
        if ($IndentationFirstLine[$i] -ne $null) {
            $Paragraph.IndentationFirstLine = $IndentationFirstLine[$i]
        }
        if ($IndentationHanging[$i] -ne $null) {
            $Paragraph.IndentationHanging = $IndentationHanging[$i]
        }
        if ($Direction[$i] -ne $null) {
            $Paragraph.Direction = $Direction[$i]
        }
    }

    if ($Supress) { return } else { return $Paragraph }
}

function Set-WordText {
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [alias ("C")] [nullable[System.Drawing.Color]]$Color,
        [alias ("S")] [nullable[double]] $FontSize,
        [alias ("FontName")] [string] $FontFamily,
        [alias ("B")] [nullable[bool]] $Bold,
        [alias ("I")] [nullable[bool]] $Italic,
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
        [Direction[]] $Direction = @(),
        [bool] $Supress = $true
    )

    $Paragraph = $Paragraph | Set-WordTextColor -Color $Color -Supress $false
    $Paragraph = $Paragraph | Set-WordTextFontSize -FontSize $FontSize -Supress $false
    $Paragraph = $Paragraph | Set-WordTextFontFamily -FontFamily $FontFamily -Supress $false
    $Paragraph = $Paragraph | Set-WordTextBold -Bold $Bold -Supress $false
    $Paragraph = $Paragraph | Set-WordTextItalic -Italic $Italic -Supress $false
}

function Set-WordTextFontSize {
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [alias ("S")] [nullable[double]] $FontSize,
        [bool] $Supress = $true
    )
    if ($Paragraph -ne $null -and $FontSize -ne $null) {
        $Paragraph = $Paragraph.FontSize($FontSize)
    }
    if ($Supress) { return } else { return $Paragraph }
}

function Set-WordTextColor {
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [alias ("C")] [nullable[System.Drawing.Color]] $Color,
        [bool] $Supress = $true
    )
    if ($Paragraph -ne $null -and $Color -ne $null) {
        $Paragraph = $Paragraph.Color($Color)
    }
    if ($Supress) { return } else { return $Paragraph }
}

function Set-WordTextBold {
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [nullable[bool]] $Bold,
        [bool] $Supress = $true
    )
    if ($Paragraph -ne $null -and $Bold -ne $null -and $Bold -eq $true) {
        $Paragraph = $Paragraph.Bold()
    }
    if ($Supress) { return } else { return $Paragraph }
}

function Set-WordTextItalic {
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [nullable[bool]] $Italic,
        [bool] $Supress = $true
    )
    if ($Paragraph -ne $null -and $Italic -ne $null -and $Italic -eq $true) {
        $Paragraph = $Paragraph.Italic()
    }
    if ($Supress) { return } else { return $Paragraph }
}

function Set-WordTextFontFamily {
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [string] $FontFamily,
        [bool] $Supress = $true
    )
    if ($Paragraph -ne $null -and $FontFamily -ne $null -and $FontFamily -ne '') {
        $Paragraph = $Paragraph.Font($FontFamily)
    }
    if ($Supress) { return } else { return $Paragraph }
}

function Set-WordTextUnderlineStyle {
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [nullable[UnderlineStyle]] $UnderlineStyle,
        [bool] $Supress = $true
    )
    if ($Paragraph -ne $null -and $UnderlineStyle -ne $null) {
        $Paragraph = $Paragraph.UnderlineStyle($UnderlineStyle)
    }
    if ($Supress) { return } else { return $Paragraph }
}

function Set-WordTextUnderlineColor {
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [nullable[System.Drawing.Color]] $UnderlineColor,
        [bool] $Supress = $true
    )
    if ($Paragraph -ne $null -and $UnderlineColor -ne $null) {
        $Paragraph = $Paragraph.UnderlineColor($UnderlineColor)
    }
    if ($Supress) { return } else { return $Paragraph }
}

function Set-WordTextSpacingAfter {
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [nullable[double]] $SpacingAfter,
        [bool] $Supress = $true
    )
    if ($Paragraph -ne $null -and $SpacingAfter -ne $null) {
        $Paragraph = $Paragraph.SpacingAfter($SpacingAfter)
    }
    if ($Supress) { return } else { return $Paragraph }
}

function Set-WordTextSpacingBefore {
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [nullable[double]] $SpacingBefore,
        [bool] $Supress = $true
    )
    if ($Paragraph -ne $null -and $SpacingBefore -ne $null) {
        $Paragraph = $Paragraph.SpacingBefore($SpacingBefore)
    }
    if ($Supress) { return } else { return $Paragraph }
}

function Set-WordTextSpacing {
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [nullable[double]] $Spacing,
        [bool] $Supress = $true
    )
    if ($Paragraph -ne $null -and $Spacing -ne $null) {
        $Paragraph = $Paragraph.Spacing($Spacing)
    }
    if ($Supress) { return } else { return $Paragraph }
}


function Set-WordTextHighlight {
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [nullable[Highlight]] $Highlight,
        [bool] $Supress = $true
    )
    if ($Paragraph -ne $null -and $Highlight -ne $null) {
        $Paragraph = $Paragraph.Highlight($Highlight)
    }
    if ($Supress) { return } else { return $Paragraph }
}

function Set-WordTextCapsStyle {
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [nullable[CapsStyle]] $CapsStyle,
        [bool] $Supress = $true
    )
    if ($Paragraph -ne $null -and $CapsStyle -ne $null) {
        $Paragraph = $Paragraph.CapsStyle($CapsStyle)
    }
    if ($Supress) { return } else { return $Paragraph }
}

function Set-WordTextStrikeThrough {
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [nullable[StrikeThrough]] $StrikeThrough,
        [bool] $Supress = $true
    )
    if ($Paragraph -ne $null -and $StrikeThrough -ne $null) {
        $Paragraph = $Paragraph.StrikeThrough($StrikeThrough)
    }
    if ($Supress) { return } else { return $Paragraph }
}

function Set-WordTextPercentageScale {
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [nullable[int]][ValidateRange( 200, 150, 100, 90, 80, 66, 50, 33)] $PercentageScale,
        [bool] $Supress = $true
    )
    if ($Paragraph -ne $null -and $PercentageScale -ne $null) {
        $Paragraph = $Paragraph.PercentageScale($PercentageScale)
    }
    if ($Supress) { return } else { return $Paragraph }
}

function Set-WordTextLanguage {
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [string]$Language,
        [bool] $Supress = $true
    )
    if ($Paragraph -ne $null -and $Language -ne $null -and $Language -ne '') {
        $Culture = [System.Globalization.CultureInfo]::GetCultureInfo($Language)
        $Paragraph = $Paragraph.Culture($Culture)
    }
    if ($Supress) { return } else { return $Paragraph }
}
