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
        [alias ("U")] [nullable[UnderlineStyle]] $UnderlineStyle,
        [alias ('UC')] [nullable[System.Drawing.Color]]$UnderlineColor,
        [alias ("SA")] [nullable[double]] $SpacingAfter,
        [alias ("SB")] [nullable[double]] $SpacingBefore ,
        [alias ("SP")] [nullable[double]] $Spacing ,
        [alias ("H")] [nullable[highlight]] $Highlight ,
        [alias ("CA")] [nullable[CapsStyle]] $CapsStyle ,
        [alias ("ST")] [nullable[StrikeThrough]] $StrikeThrough ,
        [alias ("HT")] [nullable[HeadingType]] $HeadingType ,
        [nullable[int]] $PercentageScale , # "Value must be one of the following: 200, 150, 100, 90, 80, 66, 50 or 33"
        [nullable[Misc]] $Misc ,
        [string] $Language ,
        [nullable[int]]$Kerning , # "Value must be one of the following: 8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48 or 72"
        [nullable[bool]]$Hidden ,
        [nullable[int]]$Position , #  "Value must be in the range -1585 - 1585"
        [nullable[single]] $IndentationFirstLine ,
        [nullable[single]] $IndentationHanging ,
        [nullable[Alignment]] $Alignment ,
        [nullable[Direction]] $DirectionFormatting ,
        [nullable[ShadingType]] $ShadingType ,
        [nullable[Script]] $Script ,
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
function Add-WordTableCellValue {
    [CmdletBinding()]
    param(
        [Xceed.Words.NET.InsertBeforeOrAfter] $Table,
        [int] $Row,
        [int] $Column,
        [Object] $Value,
        [int] $ParagraphNumber = 0,
        [alias ("C")] [nullable[System.Drawing.Color]]$Color,
        [alias ("S")] [nullable[double]] $FontSize,
        [alias ("FontName")] [string] $FontFamily,
        [alias ("B")] [nullable[bool]] $Bold,
        [alias ("I")] [nullable[bool]] $Italic,
        [alias ("U")] [nullable[UnderlineStyle]] $UnderlineStyle,
        [alias ('UC')] [nullable[System.Drawing.Color]]$UnderlineColor,
        [alias ("SA")] [nullable[double]] $SpacingAfter,
        [alias ("SB")] [nullable[double]] $SpacingBefore,
        [alias ("SP")] [nullable[double]] $Spacing,
        [alias ("H")] [nullable[highlight]] $Highlight,
        [alias ("CA")] [nullable[CapsStyle]] $CapsStyle,
        [alias ("ST")] [nullable[StrikeThrough]] $StrikeThrough,
        [alias ("HT")] [nullable[HeadingType]] $HeadingType,
        [nullable[int]] $PercentageScale , # "Value must be one of the following: 200, 150, 100, 90, 80, 66, 50 or 33"
        [nullable[Misc]] $Misc ,
        [string] $Language ,
        [nullable[int]]$Kerning , # "Value must be one of the following: 8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48 or 72"
        [nullable[bool]]$Hidden ,
        [nullable[int]]$Position , #  "Value must be in the range -1585 - 1585"
        [nullable[single]] $IndentationFirstLine ,
        [nullable[single]] $IndentationHanging ,
        [nullable[Alignment]] $Alignment ,
        [nullable[Direction]] $DirectionFormatting,
        [nullable[ShadingType]] $ShadingType,
        [nullable[System.Drawing.Color]]$ShadingColor,
        [nullable[Script]] $Script,
        [bool] $Supress = $false
    )
    Write-Verbose "Add-WordTableCellValue - Row: $Row Column $Column Value $Value Supress: $Supress"
    $Data = $Table.Rows[$Row].Cells[$Column].Paragraphs[$ParagraphNumber].Append($Value)
    <#
    $Data = Set-WordText -Paragraph $Data -Color $Color -FontSize $FontSize -FontFamily $FontFamily -Italic $Italic `
        -UnderlineStyle $UnderlineStyle -UnderlineColor $UnderlineColor -SpacingAfter $SpacingAfter -SpacingBefore $SpacingBefore -Spacing $Spacing `
        -Highlight $Highlight -CapsStyle $CapsStyle -StrikeThrough $StrikeThrough -HeadingType $HeadingType -PercentageScale $PercentageScale `
        -Misc $Misc -Language $Language -Kerning $Kerning -Position $Position -IndentationFirstLine $IndentationFirstLine `
        -IndentationHanging $IndentationHanging -Alignment $Alignment -Direction $DirectionFormatting -ShadingType $ShadingType -Script $Script -Supress $Supress `
        -Hidden $Hidden # -Bold $Bold
    #>
    $Data = $Data | Set-WordTextColor -Color $Color -Supress $false
    $Data = $Data | Set-WordTextFontSize -FontSize $FontSize -Supress $false
    $Data = $Data | Set-WordTextFontFamily -FontFamily $FontFamily -Supress $false
    $Data = $Data | Set-WordTextBold -Bold $Bold -Supress $false
    $Data = $Data | Set-WordTextItalic -Italic $Italic -Supress $false
    $Data = $Data | Set-WordTextUnderlineColor -UnderlineColor $UnderlineColor -Supress $false
    $Data = $Data | Set-WordTextUnderlineStyle -UnderlineStyle $UnderlineStyle -Supress $false
    $Data = $Data | Set-WordTextSpacingAfter -SpacingAfter $SpacingAfter -Supress $false
    $Data = $Data | Set-WordTextSpacingBefore -SpacingBefore $SpacingBefore -Supress $false
    $Data = $Data | Set-WordTextSpacing -Spacing $Spacing -Supress $false
    $Data = $Data | Set-WordTextHighlight -Highlight $Highlight -Supress $false
    $Data = $Data | Set-WordTextCapsStyle -CapsStyle $CapsStyle -Supress $false
    $Data = $Data | Set-WordTextStrikeThrough -StrikeThrough $StrikeThrough -Supress $false
    $Data = $Data | Set-WordTextPercentageScale -PercentageScale $PercentageScale -Supress $false
    $Data = $Data | Set-WordTextSpacing -Spacing $Spacing -Supress $false
    $Data = $Data | Set-WordTextLanguage -Language $Language -Supress $false
    $Data = $Data | Set-WordTextKerning -Kerning $Kerning -Supress $false
    $Data = $Data | Set-WordTextMisc -Misc $Misc -Supress $false
    $Data = $Data | Set-WordTextPosition -Position $Position -Supress $false
    $Data = $Data | Set-WordTextHidden -Hidden $Hidden -Supress $false
    $Data = $Data | Set-WordTextShadingType -ShadingColor $ShadingColor -ShadingType $ShadingType -Supress $false
    $Data = $Data | Set-WordTextScript -Script $Script -Supress $false
    $Data = $Data | Set-WordTextHeadingType -HeadingType $HeadingType -Supress $false
    $Data = $Data | Set-WordTextIndentationFirstLine -IndentationFirstLine $IndentationFirstLine -Supress $false
    $Data = $Data | Set-WordTextIndentationHanging -IndentationHanging $IndentationHanging -Supress $false
    $Data = $Data | Set-WordTextAlignment -Alignment $Alignment -Supress $false
    $Data = $Data | Set-WordTextDirection -Direction $Direction -Supress $false
    if ($Supress -eq $true) { return } else { return $Data }
}

function Convert-ObjectToProcess {
    [CmdletBinding()]
    param (
        $DataTable
    )
    $ObjectType = $DataTable.GetType().Name
    Write-Verbose "Convert-ObjectToProcess - GetType Before Conversion: $ObjectType"
    #$($DataTable.GetType().BaseType.Name)
    #$($DataTable.GetType().Name)
    if ($($DataTable.GetType().BaseType.Name) -eq 'Array' -and $($DataTable.GetType().Name) -eq 'Object[]') {
        Write-Verbose 'Convert-ObjectToProcess - Converting Array of Objects'
        #if ($DataTable.Count -gt 1) {
        $DataTable = $DataTable.ForEach( {[PSCustomObject]$_})
        #}

    }

    $ObjectType = $DataTable.GetType().Name
    Write-Verbose "Convert-ObjectToProcess - Table row count: $(Get-ObjectCount $DataTable)"
    Write-Verbose "Convert-ObjectToProcess - Object Type: $ObjectType"
    Write-Verbose "Convert-ObjectToProcess - BaseType.Name: $($DataTable.GetType().BaseType.Name)"
    Write-Verbose "Convert-ObjectToProcess - GetType Before Final Conversion: $ObjectType"
    If ($ObjectType -eq 'Hashtable' -or $ObjectType -eq 'OrderedDictionary' -or $ObjectType -eq 'PSCustomObject') {
        Write-Verbose 'Convert-ObjectToProcess - Skipping select for Hashtable / OrderedDictionary / PSCustomObject'
    } else {
        #if ($ObjectType -eq 'PSCustomObject') {
        #    Write-Verbose 'Convert-ObjectToProcess - Skipping all objects'
        #$DataTable = [rray] ($DataTable | Select-Object *)
        #} else {


        if ($ObjectType -eq 'Collection`1' -and $(Get-ObjectCount $DataTable) -eq 1) {
            Write-Verbose 'Convert-ObjectToProcess - Selecting all objects, returning array'
            $DataTable = [array] ($DataTable | Select-Object *)
        } else {
            Write-Verbose 'Convert-ObjectToProcess - Selecting all objects'
            $DataTable = ($DataTable | Select-Object *)
        }
        #}
    }

    $ObjectType = $DataTable.GetType().Name

    Write-Verbose "Convert-ObjectToProcess - GetType After Conversion: $ObjectType"
    return , $DataTable
}