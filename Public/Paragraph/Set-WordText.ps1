function Set-WordText {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Document.NET.InsertBeforeOrAfter[]] $Paragraph,
        [AllowNull()][string[]] $Text = @(),
        [alias ("C")] [System.Drawing.Color[]]$Color = @(),
        [alias ("S")] [double[]] $FontSize = @(),
        [alias ("FontName")] [string[]] $FontFamily = @(),
        [alias ("B")] [nullable[bool][]] $Bold = @(),
        [alias ("I")] [nullable[bool][]] $Italic = @(),
        [alias ("U")] [Xceed.Document.NET.UnderlineStyle[]] $UnderlineStyle = @(),
        [alias ('UC')] [System.Drawing.Color[]]$UnderlineColor = @(),
        [alias ("SA")] [double[]] $SpacingAfter = @(),
        [alias ("SB")] [double[]] $SpacingBefore = @(),
        [alias ("SP")] [double[]] $Spacing = @(),
        [alias ("H")] [Xceed.Document.NET.Highlight[]] $Highlight = @(),
        [alias ("CA")] [Xceed.Document.NET.CapsStyle[]] $CapsStyle = @(),
        [alias ("ST")] [Xceed.Document.NET.StrikeThrough[]] $StrikeThrough = @(),
        [alias ("HT")] [Xceed.Document.NET.HeadingType[]] $HeadingType = @(),
        [int[]] $PercentageScale = @(), # "Value must be one of the following: 200, 150, 100, 90, 80, 66, 50 or 33"
        [Xceed.Document.NET.Misc[]] $Misc = @(),
        [string[]] $Language = @(),
        [int[]]$Kerning = @(), # "Value must be one of the following: 8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48 or 72"
        [nullable[bool][]] $Hidden = @(),
        [int[]]$Position = @(), #  "Value must be in the range -1585 - 1585"
        [nullable[bool][]] $NewLine = @(),
        [switch] $KeepLinesTogether,
        [switch] $KeepWithNextParagraph,
        [single[]] $IndentationFirstLine = @(),
        [single[]] $IndentationHanging = @(),
        [nullable[Xceed.Document.NET.Alignment][]] $Alignment = @(),
        [Xceed.Document.NET.Direction[]] $Direction = @(),
        [Xceed.Document.NET.ShadingType[]] $ShadingType = @(),
        [System.Drawing.Color[]]$ShadingColor = @(),
        [Xceed.Document.NET.Script[]] $Script = @(),
        [alias ("AppendText")][Switch] $Append,
        [bool] $Supress = $false
    )
    if ($null -eq $Alignment) { $Alignment = @() }


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