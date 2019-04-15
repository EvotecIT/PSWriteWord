function Add-WordListItem {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.Container] $WordDocument,
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $List,
        [Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [InsertWhere] $InsertWhere = [InsertWhere]::AfterSelf,
        [bool] $Supress
    )
    if ($Paragraph -ne $null) {
        if ($InsertWhere -eq [InsertWhere]::AfterSelf) {
            $data = $Paragraph.InsertListAfterSelf($List)
        } elseif ($InsertWhere -eq [InsertWhere]::AfterSelf) {
            $data = $Paragraph.InsertListBeforeSelf($List)
        }
    } else {
        $data = $WordDocument.InsertList($List)
    }
    Write-Verbose "Add-WordListItem - Return Type Name: $($Data.GetType().Name) - BaseType: $($Data.GetType().BaseType)"
    if ($Supress) { return } else { $data }
}