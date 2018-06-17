function Get-ObjectTitles {
    [CmdletBinding()]
    param(
        $Object
    )
    $ArrayList = New-Object System.Collections.ArrayList
    $Titles = $Object | Get-Member | Where-Object { $_.MemberType -eq 'Property' -or $_.MemberType -eq 'NoteProperty' }
    foreach ($Title in $Titles) {
        $ArrayList.Add($Title.Name) | Out-Null
    }
    return $ArrayList
}
function Get-ObjectData {
    [CmdletBinding()]
    param(
        $Object,
        $Title,
        [switch] $DoNotAddTitles
    )
    $ArrayList = New-Object System.Collections.ArrayList
    $Values = $Object.$Title
    #Write-Color 'Get-ObjectData1: Title', ' ', $Title, ' Values: ', (Get-ObjectCount $Values) -Color Yellow, White, Green, White, Yellow
    if ((Get-ObjectCount $values) -eq 1 -and $DoNotAddTitles -eq $false) {
        $ArrayList.Add("$Title - $Values") | Out-Null
    } else {
        if ($DoNotAddTitles -eq $false) { $ArrayList.Add($Title) | Out-Null }
        foreach ($Value in $Values) {
            $ArrayList.Add("$Value") | Out-Null
        }
    }
    #Write-Color 'Get-ObjectData2: Title', ' ', $Title, ' ArrayList: ', (Get-ObjectCount $ArrayList) -Color Yellow, White, Green, White, Yellow
    return $ArrayList
}
function Get-ObjectCount {
    [CmdletBinding()]
    param(
        $Object
    )
    return $($Object | Measure-Object).Count
}
function Get-ParagraphForList {
    [CmdletBinding()]
    param(
        [Xceed.Words.NET.Container] $WordDocument,
        $ListID
    )
    $IDs = @()
    foreach ($p in $WordDocument.Paragraphs) {
        #Write-Color "testtting " -Color Yellow
        if ($p.ParagraphNumberProperties -ne $null) {
            $ListNumber = $p.ParagraphNumberProperties.LastNode.LastAttribute.Value
            if ($ListNumber -eq $ListID) {
                $IDs += $p
            }
        }
        #

        #$p.StyleName = 'Heading1'
    }
    return $Ids
}
function Get-Paragraphs {
    [CmdletBinding()]
    param(
        [Xceed.Words.NET.Container] $WordDocument
    )
    $IDs = @()
    foreach ($p in $WordDocument.Paragraphs) {
        $p
    }
    return $Ids
}