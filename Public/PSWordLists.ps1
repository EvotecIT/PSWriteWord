function Add-WordList {
    [CmdletBinding()]
    param (
        [Xceed.Words.NET.Container] $WordDocument,
        [Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [ListItemType]$ListType,
        [string[]] $ListData = $null,
        $Object = $null,
        $Supress = $true
    )
    $LevelPrimary = 0
    $LevelSecondary = 1
    $LevelThird = 2
    if ($ListData -ne $null) {
        $ListCount = ($ListData | Measure-Object).Count
        If ($ListCount -gt 0) {
            $List = $WordDocument.AddList($ListData[0], 0, $ListType)
            Write-Verbose "AddList - Name: $($List.GetType().Name) - BaseType: $($List.GetType().BaseType)"
            for ($i = 1; $i -lt $ListData.Count; $i++ ) {
                $ListItem = $WordDocument.AddListItem($List, $ListData[$i])
                Write-Verbose "AddList - Name: $($ListItem.GetType().Name) - BaseType: $($ListItem.GetType().BaseType)"
            }
            if ($Paragraph -ne $null) {
                $data = $Paragraph.InsertList($List)
            } else {
                $data = $WordDocument.InsertList($List)
            }
        }
    }
    if ($Object -ne $null) {

        $IsFirstTitle = $True
        $Titles = Get-ObjectTitles -Object $Object
        foreach ($Title in $Titles) {
            $Values = Get-ObjectData -Object $Object -Title $Title
            #$Values
            $IsFirstValue = $True
            foreach ($Value in $Values) {
                if ($IsFirstTitle -eq $True) {
                    $List = $WordDocument.AddList($Value, $LevelPrimary, $ListType)
                    Write-Verbose "AddList (Object) - Name: $($List.GetType().Name) - BaseType: $($List.GetType().BaseType)"
                } else {
                    #Write-Color 'Value IsFirstTitle ', $IsFirstTitle, ' Value IsFirstValue ', $IsFirstValue, ' Count ', $Values.Count, ' Value: ', $Value -Color Yellow, Green, Yellow, Green, White, Yellow
                    if ($IsFirstValue -eq $True) {
                        $ListItem = $WordDocument.AddListItem($List, $Value, $LevelPrimary) #> $null
                        Write-Verbose "AddList (Object) - Name: $($ListItem.GetType().Name) - BaseType: $($ListItem.GetType().BaseType)"
                    } else {
                        $ListItem = $WordDocument.AddListItem($List, $Value, $LevelSecondary) # > $null
                        Write-Verbose "AddList (Object) - Name: $($ListItem.GetType().Name) - BaseType: $($ListItem.GetType().BaseType)"
                    }
                }
                $IsFirstTitle = $false
                $IsFirstValue = $false
            }
        }
        if ($Paragraph -ne $null) {
            $data = $Paragraph.InsertList($List)
        } else {
            $data = $WordDocument.InsertList($List) #| Out-Null
        }
    }
    Write-Verbose "AddList - Name: $($data.GetType().Name) - BaseType: $($data.GetType().BaseType)"
    if ($supress -eq $false) {
        return $data
    } else {
        return
    }
}

function Convert-ListToHeadings {
    [CmdletBinding()]
    param(
        [Xceed.Words.NET.Container] $WordDocument,
        $List,
        [alias ("HT")] [HeadingType] $HeadingType = [HeadingType]::Heading1
    )
    $Headings = New-ArrayList
    $List.GetType()
    $Paragraphs = Get-WordParagraphForList $WordDocument $List.NumID
    foreach ($p in $Paragraphs) {
        $p.StyleName = $HeadingType
        Add-ToArray -List $Headings -Element $p
    }
    return $Headings
}