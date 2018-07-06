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
        [bool] $Supress = $true
    )
    Write-Verbose "Add-WordTableTitle - Title Count $($Titles.Count) Supress $Supress"

    #$Titles

    #Write-Color "Title Count $($Titles.Count) " -Color Yellow
    for ($a = 0; $a -lt $Titles.Count; $a++) {
        if ($Titles[$a] -is [string]) {
            #$Titles[$a].GetType()
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
}
function Add-WordTableCellValue {
    [CmdletBinding()]
    param(
        [Xceed.Words.NET.InsertBeforeOrAfter] $Table,
        [int] $Row,
        [int] $Column,
        [string] $Value,
        [int] $Paragraph = 0,
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
        [nullable[Script]] $Script,
        [bool] $Supress = $true
    )
    Write-Verbose "Add-WordTableCellValue - Row: $Row Column $Column Value $Value Supress: $Supress"
    #$bold.GetType()
    Write-Verbose "Add-WordTableCellValue - Bold $Bold"
    $Data = $Table.Rows[$Row].Cells[$Column].Paragraphs[$Paragraph].Append($Value)
    $Data = Set-WordText -Paragraph $Data -Color $Color -FontSize $FontSize -FontFamily $FontFamily -Italic $Italic `
        -UnderlineStyle $UnderlineStyle -UnderlineColor $UnderlineColor -SpacingAfter $SpacingAfter -SpacingBefore $SpacingBefore -Spacing $Spacing `
        -Highlight $Highlight -CapsStyle $CapsStyle -StrikeThrough $StrikeThrough -HeadingType $HeadingType -PercentageScale $PercentageScale `
        -Misc $Misc -Language $Language -Kerning $Kerning -Position $Position -IndentationFirstLine $IndentationFirstLine `
        -IndentationHanging $IndentationHanging -Alignment $Alignment -Direction $DirectionFormatting -ShadingType $ShadingType -Script $Script -Supress $Supress `
        -Hidden $Hidden # -Bold $Bold

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