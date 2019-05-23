function New-WordList {
    [CmdletBinding()]
    param(
        [ScriptBlock] $ListItems,
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.Container] $WordDocument,
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [Xceed.Words.NET.InsertBeforeOrAfter] $InsertWhere = [Xceed.Words.NET.InsertBeforeOrAfter]::AfterSelf,
        [int] $BehaviourOption = 0,
        [alias('ListType')][Xceed.Words.NET.ListItemType] $Type = [Xceed.Words.NET.ListItemType]::Bulleted,
        [bool] $Supress = $true
    )

    if ($ListItems) {
        $Parameters = Invoke-Command -ScriptBlock $ListItems

        $List = $null
        foreach ($Item in $Parameters) {
            if ($null -eq $List) {
                $List = $WordDocument.AddList($Item.Text, $Item.Level, $Type, $Item.StartNumber, $Item.TrackChanges, $Item.ContinueNumbering)
                $Paragraph = $List.Items[$List.Items.Count - 1]
            } else {
                $List = $WordDocument.AddListItem($List, $Item.Text, $Item.Level, $Type, $Item.StartNumber, $Item.TrackChanges, $Item.ContinueNumbering)
                $Paragraph = $List.Items[$List.Items.Count - 1]
            }
        }
        Add-WordListItem -WordDocument $WordDocument -List $List -Supress $true
        if (-not $Supress) {
            $List
        }
    }
}