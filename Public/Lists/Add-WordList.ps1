function Add-WordList {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.Container] $WordDocument,
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [ListItemType]$ListType = [ListItemType]::Bulleted,
        [alias('DataTable')][object] $ListData = $null,
        [InsertWhere] $InsertWhere = [InsertWhere]::AfterSelf,
        $BehaviourOption = 0,
        $ListLevels = @(),
        [bool] $Supress = $false
    )
    $List = $null
    if ($ListData -eq $null) { return }

    $Type = Get-ObjectType $ListData #-Verbose
    if ($Type.ObjectTypeName -match 'bool|byte|char|datetime|decimal|double|ExcelHyperLink|float|int|long|sbyte|short|string|timespan|uint|ulong|URI|ushort') {
        $ObjectType = $Type.ObjectTypeName
    } elseif ($Type.ObjectTypeInsiderName -ne '') {
        $ObjectType = $Type.ObjectTypeInsiderName
    } else {
        $ObjectType = $Type.ObjectTypeName
    }

    if ($ObjectType -notmatch 'bool|byte|char|datetime|decimal|double|ExcelHyperLink|float|int|long|sbyte|short|string|timespan|uint|ulong|URI|ushort' -and
        $ObjectType -ne 'PSCustomObject' -and $ObjectType -ne 'Hashtable' -and $ObjectType -ne 'OrderedDictionary') {
        $ListData = Convert-ObjectToProcess -DataTable $ListData
        $Type = Get-ObjectType $ListData
        if ($Type.ObjectTypeInsiderName -ne '') { $ObjectType = $Type.ObjectTypeInsiderName } else { $ObjectType = $Type.ObjectTypeName}
        Write-Verbose "Add-WordList - Outside Object BaseName: $($ListData.GetType().BaseType) Name: $($ListData.GetType().Name)"
        Write-Verbose "Add-WordList - Insider Object Name: $ObjectType"
    }

    if ($ObjectType -match 'bool|byte|char|datetime|decimal|double|ExcelHyperLink|float|int|long|sbyte|short|string|timespan|uint|ulong|URI|ushort') {
        Write-Verbose 'Add-WordList - Option 1 - Detected singular type inside array'
        $Counter = 0;
        foreach ($Value in $ListData) {
            if ($ListLevels -eq $null) {
                $List = New-WordListItem -WordDocument $WordDocument -List $List -ListType $ListType -ListValue $Value -ListLevel 0
                Write-Verbose "AddList - ListItemType Name: $($List.GetType().Name) - BaseType: $($List.GetType().BaseType)"
            } else {
                $List = New-WordListItem -WordDocument $WordDocument -List $List -ListType $ListType -ListValue $Value -ListLevel $ListLevels[$Counter]
                $Counter++
            }
        }
    } elseif ($ObjectType -eq 'Hashtable' -or $ObjectType -eq 'OrderedDictionary') {
        Write-Verbose "Add-WordList - Option 2 - Detected $ObjectType"
        foreach ($Object in $ListData) {
            foreach ($O in $Object.GetEnumerator()) {
                $TextMain = $($O.Name)
                $TextSub = $($O.Value)
                $List = Format-WordListItem -WordDocument $WordDocument -List $List -ListType $ListType -TextMain $TextMain -TextSub $TextSub -BehaviourOption $BehaviourOption

            }
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

    if ($supress -eq $false) {
        return $data
    } else {
        return
    }
}