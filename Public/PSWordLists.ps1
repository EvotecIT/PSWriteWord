function Add-List {
    [CmdletBinding()]
    param (
        [Xceed.Words.NET.Container] $WordDocument,
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
            for ($i = 1; $i -lt $ListData.Count; $i++ ) {
                $WordDocument.AddListItem($List, $ListData[$i]) > $null
            }
            $data = $WordDocument.InsertList($List)
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
                } else {
                    #Write-Color 'Value IsFirstTitle ', $IsFirstTitle, ' Value IsFirstValue ', $IsFirstValue, ' Count ', $Values.Count, ' Value: ', $Value -Color Yellow, Green, Yellow, Green, White, Yellow
                    if ($IsFirstValue -eq $True) {
                        $WordDocument.AddListItem($List, $Value, $LevelPrimary) > $null
                    } else {
                        $WordDocument.AddListItem($List, $Value, $LevelSecondary) > $null
                    }
                }
                $IsFirstTitle = $false
                $IsFirstValue = $false
            }
        }
        $data = $WordDocument.InsertList($List) #| Out-Null
    }

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
    $Paragraphs = Get-ParagraphForList $WordDocument $List.NumID
    foreach ($p in $Paragraphs) {
        $p.StyleName = $HeadingType
        Add-ToArray -List $Headings -Element $p
    }
    return $Headings
}