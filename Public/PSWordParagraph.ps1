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

        if ($UnderlineStyle[$i] -ne $null) {
            $Paragraph = $Paragraph.UnderlineStyle($UnderlineStyle[$i])
        }
        if ($UnderlineColor[$i] -ne $null) {
            $Paragraph = $Paragraph.UnderlineColor($UnderlineColor[$i])
        }
        if ($SpacingAfter[$i] -ne $null) {
            $Paragraph = $Paragraph.SpacingAfter($SpacingAfter[$i])
        }
        if ($SpacingBefore[$i] -ne $null) {
            $Paragraph = $Paragraph.SpacingBefore($SpacingBefore[$i])
        }
        if ($SpacingBefore[$i] -ne $null) {
            $Paragraph = $Paragraph.SpacingBefore($SpacingBefore[$i])
        }
        if ($Spacing[$i] -ne $null) {
            $Paragraph = $Paragraph.Spacing($Spacing[$i])
        }
        if ($Highlight[$i] -ne $null) {
            $Paragraph = $Paragraph.Highlight($Highlight[$i])
        }
        if ($CapsStyle[$i] -ne $null) {
            $Paragraph = $Paragraph.CapsStyle($CapsStyle[$i])
        }
        if ($StrikeThrough[$i] -ne $null) {
            $Paragraph = $Paragraph.StrikeThrough($StrikeThrough[$i])
        }
        if ($PercentageScale[$i] -ne $null) {
            $Paragraph = $Paragraph.PercentageScale($PercentageScale[$i])
        }
        if ($Language[$i] -ne $null) {
            Write-Verbose "Add-WriteText - Setting language $($Language[$i])"
            $Culture = [System.Globalization.CultureInfo]::GetCultureInfo($Language[$i])
            $Paragraph = $Paragraph.Culture($Culture)
        }
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