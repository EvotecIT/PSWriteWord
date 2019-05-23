function Add-WordTableTitle {
    [CmdletBinding()]
    param(
        [Xceed.Words.NET.InsertBeforeOrAfter] $Table,
        [string[]]$Titles,
        [int] $MaximumColumns,
        [alias ("C")] [nullable[System.Drawing.Color]]$Color,
        [alias ("S")] [nullable[double]] $FontSize,
        [alias ("FontName")] [string] $FontFamily,
        [alias ("B")] [nullable[bool]] $Bold,
        [alias ("I")] [nullable[bool]] $Italic,
        [alias ("U")] [nullable[Xceed.Words.NET.UnderlineStyle]] $UnderlineStyle,
        [alias ('UC')] [nullable[System.Drawing.Color]]$UnderlineColor,
        [alias ("SA")] [nullable[double]] $SpacingAfter,
        [alias ("SB")] [nullable[double]] $SpacingBefore ,
        [alias ("SP")] [nullable[double]] $Spacing ,
        [alias ("H")] [nullable[Xceed.Words.NET.highlight]] $Highlight ,
        [alias ("CA")] [nullable[Xceed.Words.NET.CapsStyle]] $CapsStyle ,
        [alias ("ST")] [nullable[Xceed.Words.NET.StrikeThrough]] $StrikeThrough ,
        [alias ("HT")] [nullable[Xceed.Words.NET.HeadingType]] $HeadingType ,
        [nullable[int]] $PercentageScale , # "Value must be one of the following: 200, 150, 100, 90, 80, 66, 50 or 33"
        [nullable[Xceed.Words.NET.Misc]] $Misc ,
        [string] $Language ,
        [nullable[int]]$Kerning , # "Value must be one of the following: 8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48 or 72"
        [nullable[bool]]$Hidden ,
        [nullable[int]]$Position , #  "Value must be in the range -1585 - 1585"
        [nullable[single]] $IndentationFirstLine ,
        [nullable[single]] $IndentationHanging ,
        [nullable[Xceed.Words.NET.Alignment]] $Alignment ,
        [nullable[Xceed.Words.NET.Direction]] $DirectionFormatting ,
        [nullable[Xceed.Words.NET.ShadingType]] $ShadingType ,
        [nullable[Xceed.Words.NET.Script]] $Script ,
        [bool] $Supress = $false
    )
    Write-Verbose "Add-WordTableTitle - Title Count $($Titles.Count) Supress $Supress"

    for ($a = 0; $a -lt $Titles.Count; $a++) {
        if ($Titles[$a] -is [string]) {
            $ColumnName = $Titles[$a]
        } else {
            $ColumnName = $Titles[$a].Name
        }
        Write-Verbose "Add-WordTableTitle - Column Name: $ColumnName Supress $Supress"
        Write-Verbose "Add-WordTableTitle - Bold $Bold"
        Add-WordTableCellValue -Table $Table `
            -Row 0 `
            -Column $a `
            -Value $ColumnName `
            -Color $Color -FontSize $FontSize -FontFamily $FontFamily -Bold $Bold -Italic $Italic `
            -UnderlineStyle $UnderlineStyle -UnderlineColor $UnderlineColor -SpacingAfter $SpacingAfter -SpacingBefore $SpacingBefore -Spacing $Spacing `
            -Highlight $Highlight -CapsStyle $CapsStyle -StrikeThrough $StrikeThrough -HeadingType $HeadingType -PercentageScale $PercentageScale `
            -Misc $Misc -Language $Language -Kerning $Kerning -Hidden $Hidden -Position $Position -IndentationFirstLine $IndentationFirstLine `
            -IndentationHanging $IndentationHanging -Alignment $Alignment -DirectionFormatting $DirectionFormatting -ShadingType $ShadingType -Script $Script `
            -Supress $Supress
        if ($a -eq $($MaximumColumns - 1)) {
            break;
        }
    }
    if ($Supress) { return } else { return $Table }
}

