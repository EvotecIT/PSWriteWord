function Add-WordList {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.Container] $WordDocument,
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [ListItemType]$ListType = [ListItemType]::Bulleted,
        [alias('DataTable')][object] $ListData = $null,
        [InsertWhere] $InsertWhere = [InsertWhere]::AfterSelf,
        $BehaviourOption = 0,
        [bool] $Supress = $false
    )
    $List = $null
    $ObjectType = Get-ObjectTypeInside $ListData
    if ($ObjectType -eq $null) { return }
    Write-Verbose "Add-WordList - Outside Object BaseName: $($ListData.GetType().BaseType) Name: $($ListData.GetType().Name)"
    Write-Verbose "Add-WordList - Insider Object Name: $ObjectType"

    if ($ObjectType -ne 'string' -and $ObjectType -ne 'PSCustomObject' -and $ObjectType -ne $ObjectType -ne 'Hashtable' -and $ObjectType -ne 'OrderedDictionary') {
        $ListData = Convert-ObjectToProcess -DataTable $ListData
        $ObjectType = Get-ObjectTypeInside $ListData
        Write-Verbose "Add-WordList - Outside Object BaseName: $($ListData.GetType().BaseType) Name: $($ListData.GetType().Name)"
        Write-Verbose "Add-WordList - Insider Object Name: $ObjectType"
    }

    if ($ListData -ne $null) {
        if ($ObjectType -eq 'string') {
            Write-Verbose 'Add-WordList - Option 1 - Detected string type inside array'
            foreach ($Value in $ListData) {
                $List = New-WordListItem -WordDocument $WordDocument -List $List -ListType $ListType -ListValue $Value
                Write-Verbose "AddList - ListItemType Name: $($List.GetType().Name) - BaseType: $($List.GetType().BaseType)"
            }
        } elseif ($ObjectType -eq 'Hashtable' -or $ObjectType -eq 'OrderedDictionary') {
            Write-Verbose "Add-WordList - Option 2 - Detected $ObjectType"
            foreach ($Object in $ListData) {
                foreach ($O in $Object.GetEnumerator()) {
                    $TextMain = $($O.Name)
                    $TextSub = $($O.Value)
                    $List = Format-WordListItem -WordDocument $WordDocument -List $List -ListType $ListType -TextMain $TextMain -TextSub $TextSub -BehaviourOption $BehaviourOption

                }
                <## Working alternative
                foreach ($O in $Object.Keys) {
                    Write-Verbose "Add-WordList - 2. This is Name: $O With Value $($Object.$O) "
                }
                #>
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
    }
    if ($supress -eq $false) {
        return $data
    } else {
        return
    }
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
        [alias('ListType')][ListItemType] $ListItemType,
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
        [alias ("HT")] [HeadingType] $HeadingType = [HeadingType]::Heading1
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
    return $ParagraphsWithHeadings
}