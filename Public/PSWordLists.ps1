function Add-List {
    param (
        $WordDocument,
        [ListItemType]$ListType,
        [string[]] $ListData = $null,
        $Object = $null
    )
    $LevelPrimary = 0
    $LevelSecondary = 1
    $LevelThird = 2
    if ($ListData -ne $null) {
        $ListCount = ($ListData | Measure-Object).Count
        If ($ListCount -gt 0) {
            $List = $WordDocument.AddList($ListData[0], 0, $ListType)
            for ($i = 1; $i -lt $ListData.Count; $i++ ) {
                $WordDocument.AddListItem($List, $ListData[$i]) | Out-Null
            }
            $WordDocument.InsertList($List) | Out-Null
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
                        $WordDocument.AddListItem($List, $Value, $LevelPrimary) | Out-Null
                    } else {
                        $WordDocument.AddListItem($List, $Value, $LevelSecondary) | Out-Null
                    }
                }
                $IsFirstTitle = $false
                $IsFirstValue = $false
            }
        }
        $WordDocument.InsertList($List) | Out-Null
    }


    <#
        foreach ($item in $HashData.GetEnumerator()) {
            #$item.Key
            #$item.value
            $entry = "$($item.Key) - $($item.Value)"
            if ($count -eq 0) {
                $List = $WordDocument.AddList($entry, 0, $ListType)
            } else {
                $WordDocument.AddListItem($List, $entry) | Out-Null
            }

            $count++
        }
          $WordDocument.InsertList($List) | Out-Null
          #>
}