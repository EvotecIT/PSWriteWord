function Add-WordList {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.Container] $WordDocument,
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [ListItemType]$ListType = [ListItemType]::Bulleted,
        [alias('DataTable')][object] $ListData = $null,
        [InsertWhere] $InsertWhere = [InsertWhere]::AfterSelf,
        [bool] $Supress = $false
    )
    $List = $null
    $ObjectType = Get-ObjectTypeInside $ListData
    if ($ObjectType -eq $null) { return }
    Write-Verbose "Add-WordList - Outside Object BaseName: $($ListData.GetType().BaseType) Name: $($ListData.GetType().Name)"
    Write-Verbose "Add-WordList - Insider Object Name: $ObjectType"

    if ($ListData -ne $null) {
        if ($ObjectType -eq 'string') {
            Write-Verbose 'Add-WordList - Option 1 - Detected string type inside array'
            foreach ($Value in $ListData) {
                $List = New-WordListItem -WordDocument $WordDocument -List $List -ListType $ListType -ListValue $Value
                Write-Verbose "AddList - ListItemType Name: $($List.GetType().Name) - BaseType: $($List.GetType().BaseType)"
            }
        } elseif ($ObjectType -eq 'Hashtable' -or $ObjectType -eq 'OrderedDictionary') {
            Write-Verbose "Add-WordList - Option 2 - Detected $ObjectType"
            $IsFirstValue = $True
            foreach ($Object in $ListData) {
                $IsFirstValue = $True
                foreach ($O in $Object.GetEnumerator()) {
                    #Write-Verbose "Add-WordList - 1. This is Name: $($O.Name) With Value $($O.Value) "
                    $Value = "$($O.Name) $($O.Value)"
                    Write-Verbose "Add-WordList - This is Name: $($O.Name) With Value $($O.Value) - Proposed Text: $Value "
                    if ($IsFirstValue -eq $True) {
                        $List = New-WordListItem -WordDocument $WordDocument -List $List -ListLevel 0 -ListType $ListType -ListValue $Value
                    } else {
                        $List = New-WordListItem -WordDocument $WordDocument -List $List -ListLevel 1 -ListType $ListType -ListValue $Value
                    }

                    $IsFirstValue = $false
                }
                <## Working alternative
                foreach ($O in $Object.Keys) {
                    Write-Verbose "Add-WordList - 2. This is Name: $O With Value $($Object.$O) "
                }
                #>
            }
        } elseif ($ObjectType -eq 'PSCustomObject') {
            Write-Verbose "Add-WordList - Option 3 - Detected $ObjectType"
            $IsFirstTitle = $True
            $IsFirstValue = $True
            foreach ($Object in $ListData) {
                $Titles = Get-ObjectTitles -Object $Object
                foreach ($T in $Titles) {
                    $Value = "$T - $($Object.$T)"
                    Write-Verbose "Add-WordList - This is Name: $($O.Name) With Value $($O.Value) - Proposed Text: $Value "
                    $List = New-WordListItem -WordDocument $WordDocument -List $List -ListLevel 0 -ListItemType $ListType -ListValue $Value
                }
            }
        } else {
            Write-Verbose "Add-WordList - Option 4 - Detected $ObjectType"
            $ListData = Convert-ObjectToProcess -DataTable $ListData
            $IsFirstTitle = $True
            $IsFirstValue = $True
            foreach ($Object in $ListData) {
                $Titles = Get-ObjectTitles -Object $Object
                foreach ($T in $Titles) {
                    $Value = "$T - $($Object.$T)"
                    Write-Verbose "Add-WordList - This is Name: $($O.Name) With Value $($O.Value) - Proposed Text: $Value "
                    $List = New-WordListItem -WordDocument $WordDocument -List $List -ListLevel 0 -ListItemType $ListType -ListValue $Value
                }
            }
            #throw "$ObjectType is not supported"
        }
        $Data = Add-WordListItem -WordDocument $WordDocument -List $List -Paragraph $Paragraph -Supress $Supress
    }
    if ($supress -eq $false) {
        return $data
    } else {
        return
    }
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
        [alias('Value')]$ListValue,
        [nullable[int]] $StartNumber,
        [bool]$TrackChanges = $false,
        [bool]$ContinueNumbering = $false,
        [bool]$Supress = $false
    )
    if ($List -eq $null) {
        $List = $WordDocument.AddList($ListValue, $ListLevel, $ListItemType, $StartNumber, $TrackChanges, $ContinueNumbering)
        $Paragraph = $List.Items[$List.Items.Count - 1]
    } else {
        $List = $WordDocument.AddListItem($List, $ListValue, $ListLevel, $ListItemType, $StartNumber, $TrackChanges, $ContinueNumbering)
        $Paragraph = $List.Items[$List.Items.Count - 1]
    }
    Write-Verbose "Add-WordListItem - ListType Value: $ListValue Name: $($List.GetType().Name) - BaseType: $($List.GetType().BaseType)"
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