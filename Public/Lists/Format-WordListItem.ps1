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