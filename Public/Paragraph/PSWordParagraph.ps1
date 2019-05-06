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