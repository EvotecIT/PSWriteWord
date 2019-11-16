function Add-WordListItem {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Document.NET.Container] $WordDocument,
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Document.NET.InsertBeforeOrAfter] $List,
        [Xceed.Document.NET.InsertBeforeOrAfter] $Paragraph,
        [Xceed.Document.NET.InsertBeforeOrAfter] $InsertWhere = [Xceed.Document.NET.InsertBeforeOrAfter]::AfterSelf,
        [bool] $Supress
    )
    if ($null -ne $Paragraph) {
        if ($InsertWhere -eq [Xceed.Document.NET.InsertBeforeOrAfter]::AfterSelf) {
            $data = $Paragraph.InsertListAfterSelf($List)
        } elseif ($InsertWhere -eq [Xceed.Document.NET.InsertBeforeOrAfter]::AfterSelf) {
            $data = $Paragraph.InsertListBeforeSelf($List)
        }
    } else {
        $data = $WordDocument.InsertList($List)
    }
    #Write-Verbose "Add-WordListItem - Return Type Name: $($Data.GetType().Name) - BaseType: $($Data.GetType().BaseType)"
    if ($Supress) { return } else { $data }
}