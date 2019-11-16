function New-WordListItemInternal {
    [CmdletBinding()]
    param ([parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Document.NET.Container] $WordDocument,
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Document.NET.InsertBeforeOrAfter] $List,
        [alias('Level')] [ValidateRange(0, 8)] [int] $ListLevel,
        [alias('ListType')][Xceed.Document.NET.ListItemType] $ListItemType = [Xceed.Document.NET.ListItemType]::Bulleted,
        [alias('Value', 'ListValue')]$Text,
        [nullable[int]] $StartNumber,
        [bool]$TrackChanges = $false,
        [bool]$ContinueNumbering = $false,
        [bool]$Supress = $false)
    if ($null -eq $List) {
        $List = $WordDocument.AddList($Text, $ListLevel, $ListItemType, $StartNumber, $TrackChanges, $ContinueNumbering)
    } else {
        $List = $WordDocument.AddListItem($List, $Text, $ListLevel, $ListItemType, $StartNumber, $TrackChanges, $ContinueNumbering)
    }
    $null = $List.Items[$List.Items.Count - 1]
    #$Paragraph = $List.Items[$List.Items.Count - 1]
    Write-Verbose "Add-WordListItem - ListType Value: $Text Name: $($List.GetType().Name) - BaseType: $($List.GetType().BaseType)"
    return $List
}