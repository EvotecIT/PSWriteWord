function New-DocWordList {
    [CmdletBinding()]
    param(
        [Xceed.Document.NET.Container] $WordDocument,
        [PSCustomObject] $Parameters
    )
    $List = $null
    foreach ($Item in $Parameters.ListItems) {
        if ($null -eq $List) {
            $List = $WordDocument.AddList($Item.Text, $Item.Level, $Parameters.Type, $Item.StartNumber, $Item.TrackChanges, $Item.ContinueNumbering)
            #$Paragraph = $List.Items[$List.Items.Count - 1]
        } else {
            $List = $WordDocument.AddListItem($List, $Item.Text, $Item.Level, $Parameters.Type, $Item.StartNumber, $Item.TrackChanges, $Item.ContinueNumbering)
            #$Paragraph = $List.Items[$List.Items.Count - 1]
        }
    }
    $null = Add-WordListItem -WordDocument $WordDocument -List $List
}