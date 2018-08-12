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
        [alias ("B")] [nullable[bool][]] $Bold = @(),
        [alias ("I")] [nullable[bool][]] $Italic = @(),
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
        [Misc[]] $Misc = @(),
        [string[]] $Language = @(),
        [int[]]$Kerning = @(), # "Value must be one of the following: 8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48 or 72"
        [nullable[bool][]]$Hidden = @(),
        [int[]]$Position = @(), #  "Value must be in the range -1585 - 1585"
        [nullable[bool][]]$NewLine = @(),
        # [switch] $KeepLinesTogether, # not done
        # [switch] $KeepWithNextParagraph, # not done
        [single[]] $IndentationFirstLine = @(),
        [single[]] $IndentationHanging = @(),
        [Alignment[]] $Alignment = @(),
        [Direction[]] $Direction = @(),
        [ShadingType[]] $ShadingType = @(),
        [System.Drawing.Color[]]$ShadingColor = @(),
        [Script[]] $Script = @(),
        [Switch] $ContinueFormatting,
        [alias ("Append")][Switch] $AppendToExistingParagraph,
        [bool] $Supress = $false
    )
    if ($Alignment -eq $null) { $Alignment = @() }
    if ($Text.Count -eq 0) { return }

    if ($Paragraph -ne $null) {
        if (-not $AppendToExistingParagraph) {
            $NewParagraph = $WordDocument.InsertParagraph()
            $Paragraph = $Paragraph.InsertParagraphAfterSelf($NewParagraph)
        }
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
                    $Paragraph = $Paragraph.InsertParagraphAfterSelf($Paragraph)
                } else {
                    $Paragraph = $WordDocument.InsertParagraph()
                }
            }
            $Paragraph = $Paragraph.Append($Text[$i])
        } else {
            $Paragraph = $Paragraph.Append($Text[$i])
        }

        if ($ContinueFormatting -eq $true) {
            Write-Verbose "Add-WordText - ContinueFormatting: $ContinueFormatting Text Count: $($Text.Count)"
            $Formatting = Set-WordContinueFormatting -Count $Text.Count `
                -Color $Color `
                -FontSize $FontSize `
                -FontFamily $FontFamily `
                -Bold $Bold `
                -Italic $Italic `
                -UnderlineStyle $UnderlineStyle `
                -UnderlineColor $UnderlineColor `
                -SpacingAfter $SpacingAfter `
                -SpacingBefore $SpacingBefore `
                -Spacing $Spacing `
                -Highlight $Highlight `
                -CapsStyle $CapsStyle `
                -StrikeThrough $StrikeThrough `
                -HeadingType $HeadingType `
                -PercentageScale $PercentageScale `
                -Misc $Misc `
                -Language $Language `
                -Kerning $Kerning `
                -Hidden $Hidden `
                -Position $Position `
                -IndentationFirstLine $IndentationFirstLine `
                -IndentationHanging $IndentationHanging `
                -Alignment $Alignment `
                -ShadingType $ShadingType `
                -Script $Script

            $Color = $Formatting[0]
            $FontSize = $Formatting[1]
            $FontFamily = $Formatting[2]
            $Bold = $Formatting[3]
            $Italic = $Formatting[4]
            $UnderlineStyle = $Formatting[5]
            $UnderlineColor = $Formatting[6]
            $SpacingAfter = $Formatting[7]
            $SpacingBefore = $Formatting[8]
            $Spacing = $Formatting[9]
            $Highlight = $Formatting[10]
            $CapsStyle = $Formatting[11]
            $StrikeThrough = $Formatting[12]
            $HeadingType = $Formatting[13]
            $PercentageScale = $Formatting[14]
            $Misc = $Formatting[15]
            $Language = $Formatting[16]
            $Kerning = $Formatting[17]
            $Hidden = $Formatting[18]
            $Position = $Formatting[19]
            $IndentationFirstLine = $Formatting[20]
            $IndentationHanging = $Formatting[21]
            $Alignment = $Formatting[22]
            #$DirectionFormatting = $Formatting[23]
            $ShadingType = $Formatting[24]
            $Script = $Formatting[25]
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
        $Paragraph = $Paragraph | Set-WordTextKerning -Kerning $Kerning[$i] -Supress $false
        $Paragraph = $Paragraph | Set-WordTextMisc -Misc $Misc[$i] -Supress $false
        $Paragraph = $Paragraph | Set-WordTextPosition -Position $Position[$i] -Supress $false
        $Paragraph = $Paragraph | Set-WordTextHidden -Hidden $Hidden[$i] -Supress $false
        $Paragraph = $Paragraph | Set-WordTextShadingType -ShadingColor $ShadingColor[$i] -ShadingType $ShadingType[$i] -Supress $false
        $Paragraph = $Paragraph | Set-WordTextScript -Script $Script[$i] -Supress $false
        $Paragraph = $Paragraph | Set-WordTextHeadingType -HeadingType $HeadingType[$i] -Supress $false
        $Paragraph = $Paragraph | Set-WordTextIndentationFirstLine -IndentationFirstLine $IndentationFirstLine[$i] -Supress $false
        $Paragraph = $Paragraph | Set-WordTextIndentationHanging -IndentationHanging $IndentationHanging[$i] -Supress $false
        $Paragraph = $Paragraph | Set-WordTextAlignment -Alignment $Alignment[$i] -Supress $false
        $Paragraph = $Paragraph | Set-WordTextDirection -Direction $Direction[$i] -Supress $false
    }

    if ($Supress) { return } else { return $Paragraph }
}

function Set-WordText {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter[]] $Paragraph,
        [AllowNull()][string[]] $Text = @(),
        [alias ("C")] [System.Drawing.Color[]]$Color = @(),
        [alias ("S")] [double[]] $FontSize = @(),
        [alias ("FontName")] [string[]] $FontFamily = @(),
        [alias ("B")] [nullable[bool][]] $Bold = @(),
        [alias ("I")] [nullable[bool][]] $Italic = @(),
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
        [Misc[]] $Misc = @(),
        [string[]] $Language = @(),
        [int[]]$Kerning = @(), # "Value must be one of the following: 8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48 or 72"
        [nullable[bool][]] $Hidden = @(),
        [int[]]$Position = @(), #  "Value must be in the range -1585 - 1585"
        [nullable[bool][]] $NewLine = @(),
        [switch] $KeepLinesTogether,
        [switch] $KeepWithNextParagraph,
        [single[]] $IndentationFirstLine = @(),
        [single[]] $IndentationHanging = @(),
        [nullable[Alignment][]] $Alignment = @(),
        [Direction[]] $Direction = @(),
        [ShadingType[]] $ShadingType = @(),
        [System.Drawing.Color[]]$ShadingColor = @(),
        [Script[]] $Script = @(),
        [alias ("AppendText")][Switch] $Append,
        [bool] $Supress = $false
    )
    if ($Alignment -eq $null) { $Alignment = @() }


    Write-Verbose "Set-WordText - Paragraph Count: $($Paragraph.Count)"
    for ($i = 0; $i -lt $Paragraph.Count; $i++) {
        Write-Verbose "Set-WordText - Loop: $($i)"
        Write-Verbose "Set-WordText - $($Paragraph[$i])"
        Write-Verbose "Set-WordText - $($Paragraph[$i].Text)"
        if ($null -eq $Paragraph[$i]) {
            Write-Verbose 'Set-WordText - Paragraph is null'
        } else {
            Write-Verbose 'Set-WordText - Paragraph is not null'
        }
        if ($null -eq $Color[$i]) {
            Write-Verbose 'Set-WordText - Color is null'
        } else {
            Write-Verbose 'Set-WordText - Color is not null'
        }
        $Paragraph[$i] = $Paragraph[$i] | Set-WordTextText -Text $Text[$i] -Append:$Append -Supress $false
        $Paragraph[$i] = $Paragraph[$i] | Set-WordTextColor -Color $Color[$i] -Supress $false
        $Paragraph[$i] = $Paragraph[$i] | Set-WordTextFontSize -FontSize $FontSize[$i] -Supress $false
        $Paragraph[$i] = $Paragraph[$i] | Set-WordTextFontFamily -FontFamily $FontFamily[$i] -Supress $false
        $Paragraph[$i] = $Paragraph[$i] | Set-WordTextBold -Bold $Bold[$i] -Supress $false
        $Paragraph[$i] = $Paragraph[$i] | Set-WordTextItalic -Italic $Italic[$i] -Supress $false
        $Paragraph[$i] = $Paragraph[$i] | Set-WordTextUnderlineColor -UnderlineColor $UnderlineColor[$i] -Supress $false
        $Paragraph[$i] = $Paragraph[$i] | Set-WordTextUnderlineStyle -UnderlineStyle $UnderlineStyle[$i] -Supress $false
        $Paragraph[$i] = $Paragraph[$i] | Set-WordTextSpacingAfter -SpacingAfter $SpacingAfter[$i] -Supress $false
        $Paragraph[$i] = $Paragraph[$i] | Set-WordTextSpacingBefore -SpacingBefore $SpacingBefore[$i] -Supress $false
        $Paragraph[$i] = $Paragraph[$i] | Set-WordTextSpacing -Spacing $Spacing[$i] -Supress $false
        $Paragraph[$i] = $Paragraph[$i] | Set-WordTextHighlight -Highlight $Highlight[$i] -Supress $false
        $Paragraph[$i] = $Paragraph[$i] | Set-WordTextCapsStyle -CapsStyle $CapsStyle[$i] -Supress $false
        $Paragraph[$i] = $Paragraph[$i] | Set-WordTextStrikeThrough -StrikeThrough $StrikeThrough[$i] -Supress $false
        $Paragraph[$i] = $Paragraph[$i] | Set-WordTextPercentageScale -PercentageScale $PercentageScale[$i] -Supress $false
        $Paragraph[$i] = $Paragraph[$i] | Set-WordTextSpacing -Spacing $Spacing[$i] -Supress $false
        $Paragraph[$i] = $Paragraph[$i] | Set-WordTextLanguage -Language $Language[$i] -Supress $false
        $Paragraph[$i] = $Paragraph[$i] | Set-WordTextKerning -Kerning $Kerning[$i] -Supress $false
        $Paragraph[$i] = $Paragraph[$i] | Set-WordTextMisc -Misc $Misc[$i] -Supress $false
        $Paragraph[$i] = $Paragraph[$i] | Set-WordTextPosition -Position $Position[$i] -Supress $false
        $Paragraph[$i] = $Paragraph[$i] | Set-WordTextHidden -Hidden $Hidden[$i] -Supress $false
        $Paragraph[$i] = $Paragraph[$i] | Set-WordTextShadingType -ShadingColor $ShadingColor[$i] -ShadingType $ShadingType[$i] -Supress $false
        $Paragraph[$i] = $Paragraph[$i] | Set-WordTextScript -Script $Script[$i] -Supress $false
        $Paragraph[$i] = $Paragraph[$i] | Set-WordTextHeadingType -HeadingType $HeadingType[$i] -Supress $false
        $Paragraph[$i] = $Paragraph[$i] | Set-WordTextIndentationFirstLine -IndentationFirstLine $IndentationFirstLine[$i] -Supress $false
        $Paragraph[$i] = $Paragraph[$i] | Set-WordTextIndentationHanging -IndentationHanging $IndentationHanging[$i] -Supress $false
        $Paragraph[$i] = $Paragraph[$i] | Set-WordTextAlignment -Alignment $Alignment[$i] -Supress $false
        $Paragraph[$i] = $Paragraph[$i] | Set-WordTextDirection -Direction $Direction[$i] -Supress $false
    }
}
function Remove-WordText {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [int] $Index = 0,
        [int] $Count = $($Paragraph.Text.Length),
        [bool] $TrackChanges,
        [bool] $RemoveEmptyParagraph,
        [bool] $Supress = $false
    )
    if ($Paragraph -ne $null) {
        Write-Verbose "Remove-WordText - Current text $($Paragraph.Text) "
        Write-Verbose "Remove-WordText - Removing from $Index to $Count - Paragraph Text Count: $($Paragraph.Text.Length)"
        $Paragraph.RemoveText($Index, $Count, $TrackChanges, $RemoveEmptyParagraph)
    }
    if ($Supress) { return } else { return $Paragraph }
}

function Set-WordTextText {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [alias ("S")][AllowNull()] $Text,
        [switch]$Append,
        [bool] $Supress = $false
    )
    if ($Paragraph -ne $null) {
        if ($Text -ne $null) {
            if ($Text -isnot [String]) { throw 'Invalid argument for parameter -Text.' }
            if ($Append -ne $true) { $Paragraph = Remove-WordText -Paragraph $Paragraph }
            Write-Verbose "Set-WordTextText - Appending Value $Text"
            $Paragraph = $Paragraph.Append($Text)
        }
    }
    if ($Supress) { return } else { return $Paragraph }
}

function Set-WordTextFontSize {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [alias ("S")] [nullable[double]] $FontSize,
        [bool] $Supress = $false
    )
    if ($Paragraph -ne $null -and $FontSize -ne $null) {
        $Paragraph = $Paragraph.FontSize($FontSize)
    }
    if ($Supress) { return } else { return $Paragraph }
}

function Set-WordTextColor {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [alias ("C")] [nullable[System.Drawing.Color]] $Color,
        [bool] $Supress = $false
    )
    if ($Paragraph -ne $null -and $Color -ne $null) {
        $Paragraph = $Paragraph.Color($Color)
    }
    if ($Supress) { return } else { return $Paragraph }
}

function Set-WordTextBold {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [nullable[bool]] $Bold,
        [bool] $Supress = $false
    )
    if ($Paragraph -ne $null -and $Bold -ne $null -and $Bold -eq $true) {
        $Paragraph = $Paragraph.Bold()
    }
    if ($Supress) { return } else { return $Paragraph }
}

function Set-WordTextItalic {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [nullable[bool]] $Italic,
        [bool] $Supress = $false
    )
    if ($Paragraph -ne $null -and $Italic -ne $null -and $Italic -eq $true) {
        $Paragraph = $Paragraph.Italic()
    }
    if ($Supress) { return } else { return $Paragraph }
}

function Set-WordTextFontFamily {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [string] $FontFamily,
        [bool] $Supress = $false
    )
    if ($Paragraph -ne $null -and $FontFamily -ne $null -and $FontFamily -ne '') {
        $Paragraph = $Paragraph.Font($FontFamily)
    }
    if ($Supress) { return } else { return $Paragraph }
}

function Set-WordTextUnderlineStyle {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [nullable[UnderlineStyle]] $UnderlineStyle,
        [bool] $Supress = $false
    )
    if ($Paragraph -ne $null -and $UnderlineStyle -ne $null) {
        $Paragraph = $Paragraph.UnderlineStyle($UnderlineStyle)
    }
    if ($Supress) { return } else { return $Paragraph }
}

function Set-WordTextUnderlineColor {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [nullable[System.Drawing.Color]] $UnderlineColor,
        [bool] $Supress = $false
    )
    if ($Paragraph -ne $null -and $UnderlineColor -ne $null) {
        $Paragraph = $Paragraph.UnderlineColor($UnderlineColor)
    }
    if ($Supress) { return } else { return $Paragraph }
}

function Set-WordTextSpacingAfter {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [nullable[double]] $SpacingAfter,
        [bool] $Supress = $false
    )
    if ($Paragraph -ne $null -and $SpacingAfter -ne $null) {
        $Paragraph = $Paragraph.SpacingAfter($SpacingAfter)
    }
    if ($Supress) { return } else { return $Paragraph }
}

function Set-WordTextSpacingBefore {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [nullable[double]] $SpacingBefore,
        [bool] $Supress = $false
    )
    if ($Paragraph -ne $null -and $SpacingBefore -ne $null) {
        $Paragraph = $Paragraph.SpacingBefore($SpacingBefore)
    }
    if ($Supress) { return } else { return $Paragraph }
}

function Set-WordTextSpacing {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [nullable[double]] $Spacing,
        [bool] $Supress = $false
    )
    if ($Paragraph -ne $null -and $Spacing -ne $null) {
        $Paragraph = $Paragraph.Spacing($Spacing)
    }
    if ($Supress) { return } else { return $Paragraph }
}


function Set-WordTextHighlight {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [nullable[Highlight]] $Highlight,
        [bool] $Supress = $false
    )
    if ($Paragraph -ne $null -and $Highlight -ne $null) {
        $Paragraph = $Paragraph.Highlight($Highlight)
    }
    if ($Supress) { return } else { return $Paragraph }
}

function Set-WordTextCapsStyle {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [nullable[CapsStyle]] $CapsStyle,
        [bool] $Supress = $false
    )
    if ($Paragraph -ne $null -and $CapsStyle -ne $null) {
        $Paragraph = $Paragraph.CapsStyle($CapsStyle)
    }
    if ($Supress) { return } else { return $Paragraph }
}

function Set-WordTextStrikeThrough {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [nullable[StrikeThrough]] $StrikeThrough,
        [bool] $Supress = $false
    )
    if ($Paragraph -ne $null -and $StrikeThrough -ne $null) {
        $Paragraph = $Paragraph.StrikeThrough($StrikeThrough)
    }
    if ($Supress) { return } else { return $Paragraph }
}
function Set-WordTextShadingType {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [nullable[ShadingType]] $ShadingType,
        [nullable[System.Drawing.Color]] $ShadingColor,
        [bool] $Supress = $false
    )
    if ($Paragraph -ne $null -and $ShadingType -ne $null -and $ShadingColor -ne $null) {
        $Paragraph = $Paragraph.Shading($ShadingColor, $ShadingType)
    }
    if ($Supress) { return } else { return $Paragraph }
}

function Set-WordTextPercentageScale {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [nullable[int]]$PercentageScale,
        [bool] $Supress = $false
    )
    if ($Paragraph -ne $null -and $PercentageScale -ne $null) {
        $Paragraph = $Paragraph.PercentageScale($PercentageScale)
    }
    if ($Supress) { return } else { return $Paragraph }
}

function Set-WordTextLanguage {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [string]$Language,
        [bool] $Supress = $false
    )
    if ($Paragraph -ne $null -and $Language -ne $null -and $Language -ne '') {
        $Culture = [System.Globalization.CultureInfo]::GetCultureInfo($Language)
        $Paragraph = $Paragraph.Culture($Culture)
    }
    if ($Supress) { return } else { return $Paragraph }
}

function Set-WordTextKerning {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [nullable[int]] $Kerning,
        [bool] $Supress = $false
    )
    if ($Paragraph -ne $null -and $Kerning -ne $null) {
        $Paragraph = $Paragraph.Kerning($Kerning)
    }
    if ($Supress) { return } else { return $Paragraph }
}

function Set-WordTextMisc {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [nullable[Misc]] $Misc,
        [bool] $Supress = $false
    )
    if ($Paragraph -ne $null -and $Misc -ne $null) {
        $Paragraph = $Paragraph.Misc($Misc)
    }
    if ($Supress) { return } else { return $Paragraph }
}


function Set-WordTextPosition {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [nullable[int]]$Position,
        [bool] $Supress = $false
    )
    if ($Paragraph -ne $null -and $Position -ne $null) {
        $Paragraph = $Paragraph.Position($Position)
    }
    if ($Supress) { return } else { return $Paragraph }
}

function Set-WordTextHidden {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [nullable[bool]] $Hidden,
        [bool] $Supress = $false
    )
    if ($Paragraph -ne $null -and $Hidden -ne $null) {
        $Paragraph = $Paragraph.Hidden($Hidden)
    }
    if ($Supress) { return } else { return $Paragraph }
}

function Set-WordTextHeadingType {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [nullable[HeadingType]] $HeadingType,
        [bool] $Supress = $false
    )
    if ($Paragraph -ne $null -and $HeadingType -ne $null) {
        #$StyleName = [string] "$HeadingType"
        Write-Verbose "Set-WordTextHeadingType - Setting StyleName to $StyleName"
        $Paragraph.StyleName = $HeadingType
    }
    if ($Supress) { return } else { return $Paragraph }
}
function Set-WordTextIndentationFirstLine {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [nullable[single]] $IndentationFirstLine,
        [bool] $Supress = $false
    )
    if ($Paragraph -ne $null -and $IndentationFirstLine -ne $null) {
        $Paragraph.IndentationFirstLine = $IndentationFirstLine
    }
    if ($Supress) { return } else { return $Paragraph }
}
function Set-WordTextAlignment {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [nullable[Alignment]] $Alignment,
        [bool] $Supress = $false
    )
    if ($Paragraph -ne $null -and $Alignment -ne $null) {
        $Paragraph.Alignment = $Alignment
    }
    if ($Supress) { return } else { return $Paragraph }
}
function Set-WordTextIndentationHanging {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [nullable[single]] $IndentationHanging,
        [bool] $Supress = $false
    )
    if ($Paragraph -ne $null -and $IndentationHanging -ne $null) {
        $Paragraph.IndentationHanging = $IndentationHanging
    }
    if ($Supress) { return } else { return $Paragraph }
}

function Set-WordTextDirection {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [nullable[Direction]] $Direction,
        [bool] $Supress = $false
    )
    if ($Paragraph -ne $null -and $Direction -ne $null) {
        $Paragraph.Direction = $Direction
    }
    if ($Supress) { return } else { return $Paragraph }
}
function Set-WordTextScript {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [nullable[Script]] $Script,
        [bool] $Supress = $false
    )
    if ($Paragraph -ne $null -and $Script -ne $null) {
        $Paragraph = $Paragraph.Script($Script)
    }
    if ($Supress) { return } else { return $Paragraph }
}

function Remove-WordParagraph {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [bool] $TrackChanges
    )
    $Paragraph.Remove($TrackChanges)
}