function Add-WordList {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Container] $WordDocument,
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][InsertBeforeOrAfter] $Paragraph,
        [alias('ListType')][ListItemType] $Type = [ListItemType]::Bulleted,
        [alias('DataTable')][Array] $ListData,
        #[alias('Insert')][validateset('BeforeSelf', 'AfterSelf')][string] $InsertWhere = 'AfterSelf',
        [int] $BehaviourOption = 0,
        [Array] $ListLevels,
        [bool] $Supress = $false
    )
    if ($ListData.Count -gt 0) {

        if ($ListData[0].GetType() -match 'bool|byte|char|datetime|decimal|double|float|int|long|sbyte|short|string|timespan|uint|ulong|URI|ushort') {
            $Counter = 0;
            $Data = New-WordList -WordDocument $WordDocument -Type $Type {
                foreach ($Item in $ListData) {
                    if ($ListLevels) {
                        New-WordListItem -Level $ListLevels[$Counter] -Text $Item
                    } else {
                        New-WordListItem -Level 0 -Text $Item
                    }
                    $Counter++
                }

            } -Supress $Supress

        } elseif ($ListData[0] -is [System.Collections.IDictionary]) {
            $Data = New-WordList -WordDocument $WordDocument -Type $Type {
                foreach ($Object in $ListData) {
                    foreach ($O in $Object.GetEnumerator()) {
                        $TextMain = $($O.Name)
                        $TextSub = $($O.Value)

                        if ($BehaviourOption -eq 0) {
                            New-WordListItem -ListLevel 0 -ListValue $TextMain
                            foreach ($TextValue in $TextSub) {
                                New-WordListItem -ListLevel 1 -ListValue $TextValue
                            }
                        } elseif ($BehaviourOption -eq 1) {
                            $TextSub = $TextSub -Join ", "
                            $Value = "$TextMain - $TextSub"
                            New-WordListItem  -ListLevel 0  -ListValue $Value
                        }

                    }
                }
            } -Supress $supress
        } else {
            $Data = New-WordList -WordDocument $WordDocument -Type $Type {
                foreach ($Object in $ListData) {
                    $Titles = $Object.PSObject.Properties.Name
                    foreach ($Text in $Titles) {
                        $TextMain = $Text
                        $TextSub = $($Object.$Text)
                        if ($BehaviourOption -eq 0) {
                            New-WordListItem  -ListLevel 0 -ListValue $TextMain
                            foreach ($TextValue in $TextSub) {
                                New-WordListItem  -ListLevel 1 -ListValue $TextValue
                            }
                        } elseif ($BehaviourOption -eq 1) {
                            $TextSub = $TextSub -Join ", "
                            $Value = "$TextMain - $TextSub"
                            New-WordListItem -ListLevel 0  -ListValue $Value
                        }
                    }
                }
            } -Supress $Supress
        }
        if ($supress -eq $false) {
            return $Data
        } else {
            return
        }
    }

    <#
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
                $List = New-WordListItemInternal -WordDocument $WordDocument -List $List -ListType $ListType -ListValue $Value -ListLevel 0
                Write-Verbose "AddList - ListItemType Name: $($List.GetType().Name) - BaseType: $($List.GetType().BaseType)"
            } else {
                $List = New-WordListItemInternal -WordDocument $WordDocument -List $List -ListType $ListType -ListValue $Value -ListLevel $ListLevels[$Counter]
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
    #>
}