function Add-WordList {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.Container] $WordDocument,
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [ListItemType]$ListType = [ListItemType]::Bulleted,
        [alias('DataTable')][object] $ListData = $null,
        [InsertWhere] $InsertWhere = [InsertWhere]::AfterSelf,
        $BehaviourOption = 0,
        $ListLevels = @(),
        [bool] $Supress = $false
    )
    $List = $null
    if ($ListData -eq $null) { return }

    $Type = Get-ObjectType $ListData
    if ($Type.ObjectTypeInsiderName -ne '') { $ObjectType = $Type.ObjectTypeInsiderName } else { $ObjectType = $Type.ObjectTypeName}

    if ($ObjectType -ne 'string' -and $ObjectType -ne 'PSCustomObject' -and $ObjectType -ne 'Hashtable' -and $ObjectType -ne 'OrderedDictionary') {
        $ListData = Convert-ObjectToProcess -DataTable $ListData
        $Type = Get-ObjectType $ListData
        if ($Type.ObjectTypeInsiderName -ne '') { $ObjectType = $Type.ObjectTypeInsiderName } else { $ObjectType = $Type.ObjectTypeName}
        Write-Verbose "Add-WordList - Outside Object BaseName: $($ListData.GetType().BaseType) Name: $($ListData.GetType().Name)"
        Write-Verbose "Add-WordList - Insider Object Name: $ObjectType"
    }

    if ($ObjectType -eq 'string') {
        Write-Verbose 'Add-WordList - Option 1 - Detected string type inside array'
        $Counter = 0;
        foreach ($Value in $ListData) {
            if ($ListLevels -eq $null) {
                $List = New-WordListItem -WordDocument $WordDocument -List $List -ListType $ListType -ListValue $Value -ListLevel 0
                Write-Verbose "AddList - ListItemType Name: $($List.GetType().Name) - BaseType: $($List.GetType().BaseType)"
            } else {
                $List = New-WordListItem -WordDocument $WordDocument -List $List -ListType $ListType -ListValue $Value -ListLevel $ListLevels[$Counter]
                $Counter++
            }
        }
    } elseif ($ObjectType -eq 'Hashtable' -or $ObjectType -eq 'OrderedDictionary') {
        Write-Verbose "Add-WordList - Option 2 - Detected $ObjectType"
        foreach ($Object in $ListData) {
            foreach ($O in $Object.GetEnumerator()) {
                $TextMain = $($O.Name)
                $TextSub = $($O.Value)
                $List = Format-WordListItem -WordDocument $WordDocument -List $List -ListType $ListType -TextMain $TextMain -TextSub $TextSub -BehaviourOption $BehaviourOption

            }
        }
    } elseif ($ObjectType -eq 'PSCustomObject') {
        Write-Verbose "Add-WordList - Option 3 - Detected $ObjectType"
        foreach ($Object in $ListData) {
            $Titles = Get-ObjectTitles -Object $Object
            foreach ($Text in $Titles) {
                $TextMain = $Text
                $TextSub = $($Object.$Text)
                $List = Format-WordListItem -WordDocument $WordDocument -List $List -ListType $ListType -TextMain $TextMain -TextSub $TextSub -BehaviourOption $BehaviourOption
            }
        }
    } else {
        throw "$ObjectType is not supported - report for support with explanation what you need it to look like"
    }
    $Data = Add-WordListItem -WordDocument $WordDocument -List $List -Paragraph $Paragraph -Supress $Supress

    if ($supress -eq $false) {
        return $data
    } else {
        return
    }
}

function Set-WordList {
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.Container] $WordDocument,
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $List,
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
    foreach ($Data in $List.Items) {
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
    }
    if ($Supress) { return } else { return $List }
}

function Format-WordListItem {
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.Container] $WordDocument,
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $List,
        [ListItemType]$ListType = [ListItemType]::Bulleted,
        $TextMain,
        $TextSub,
        $BehaviourOption
    )

    if ($BehaviourOption -eq 0) {
        Write-Verbose "Add-WordList - This is Name: $($TextMain) With Value $TextSub - Proposed Text: $TextMain and $TextSub on separate line "
        $List = New-WordListItem -WordDocument $WordDocument -List $List -ListLevel 0 -ListItemType $ListType -ListValue $TextMain
        foreach ($TextValue in $TextSub) {
            $List = New-WordListItem -WordDocument $WordDocument -List $List -ListLevel 1 -ListItemType $ListType -ListValue $TextValue
        }
    } elseif ($BehaviourOption -eq 1) {
        $TextSub = $TextSub -Join ", "
        $Value = "$TextMain - $TextSub"
        Write-Verbose "Add-WordList - This is Name: $($TextMain) With Value $TextSub - Proposed Text: $Value "
        $List = New-WordListItem -WordDocument $WordDocument -List $List -ListLevel 0 -ListItemType $ListType -ListValue $Value
    }
    return $List
}



function Add-WordListItem {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.Container] $WordDocument,
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $List,
        [Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [InsertWhere] $InsertWhere = [InsertWhere]::AfterSelf,
        [bool] $Supress
    )
    if ($Paragraph -ne $null) {
        if ($InsertWhere -eq [InsertWhere]::AfterSelf) {
            $data = $Paragraph.InsertListAfterSelf($List)
        } elseif ($InsertWhere -eq [InsertWhere]::AfterSelf) {
            $data = $Paragraph.InsertListBeforeSelf($List)
        }
    } else {
        $data = $WordDocument.InsertList($List)
    }
    Write-Verbose "Add-WordListItem - Return Type Name: $($Data.GetType().Name) - BaseType: $($Data.GetType().BaseType)"
    if ($Supress) { return } else { $data }
}

function New-WordListItem {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.Container] $WordDocument,
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $List,
        [alias('Level')] [ValidateRange(0, 8)] [int] $ListLevel,
        [alias('ListType')][ListItemType] $ListItemType = [ListItemType]::Bulleted,
        [alias('Value', 'ListValue')]$Text,
        [nullable[int]] $StartNumber,
        [bool]$TrackChanges = $false,
        [bool]$ContinueNumbering = $false,
        [bool]$Supress = $false
    )
    if ($List -eq $null) {
        $List = $WordDocument.AddList($Text, $ListLevel, $ListItemType, $StartNumber, $TrackChanges, $ContinueNumbering)
        $Paragraph = $List.Items[$List.Items.Count - 1]
    } else {
        $List = $WordDocument.AddListItem($List, $Text, $ListLevel, $ListItemType, $StartNumber, $TrackChanges, $ContinueNumbering)
        $Paragraph = $List.Items[$List.Items.Count - 1]
    }
    Write-Verbose "Add-WordListItem - ListType Value: $Text Name: $($List.GetType().Name) - BaseType: $($List.GetType().BaseType)"
    return $List
}

function Get-WordListItemParagraph {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $List,
        [nullable[int]] $Item,
        [switch] $LastItem
    )
    if ($List -ne $null) {
        $Count = $List.Items.Count
        Write-Verbose "Get-WordListItemParagraph - List Count $Count"
        if ($LastItem) {
            Write-Verbose "Get-WordListItemParagraph - Last Element $($Count-1)"
            $Paragraph = $List.Items[$Count - 1]
        } else {
            if ($null -ne $Item -and $Item -le $Count) {
                Write-Verbose "Get-WordListItemParagraph - Returning paragraph for Item Nr: $Item"
                $Paragraph = $List.Items[$Item]
            }
        }
    }
    return $Paragraph
}

function Convert-ListToHeadings {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.Container] $WordDocument,
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $List,
        [alias ("HT")] [HeadingType] $HeadingType = [HeadingType]::Heading1,
        [bool] $Supress
    )
    $ParagraphsWithHeadings = New-ArrayList
    Write-Verbose "Convert-ListToHeadings - NumID: $($List.NumID)"
    $Paragraphs = Get-WordParagraphForList -WordDocument $WordDocument -ListID $List.NumID
    Write-Verbose "Convert-ListToHeadings - List Elements Count: $($Paragraphs.Count)"
    foreach ($p in $Paragraphs) {
        Write-Verbose "Convert-ListToHeadings - Loop: $HeadingType"
        $p.StyleName = $HeadingType
        Add-ToArray -List $ParagraphsWithHeadings -Element $p
    }
    if ($Supress) { return } else { return $ParagraphsWithHeadings }
}