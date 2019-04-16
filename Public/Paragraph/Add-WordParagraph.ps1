Function Add-WordParagraph {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.Container]$WordDocument,
        [alias('Paragraph', 'Table', 'List')][Xceed.Words.NET.InsertBeforeOrAfter] $WordObject,
        [alias('Insert')][InsertWhere] $InsertWhere = [InsertWhere]::AfterSelf,
        #[bool] $TrackChanges,
        [bool] $Supress = $false
    )
    $NewParagraph = $WordDocument.InsertParagraph()
    if ($WordObject -ne $null) {
        if ($InsertWhere -eq [InsertWhere]::AfterSelf) {
            $NewParagraph = $WordObject.InsertParagraphAfterSelf($NewParagraph)
        } elseif ($InsertWhere -eq [InsertWhere]::BeforeSelf) {
            $NewParagraph = $WordObject.InsertParagraphBeforeSelf($NewParagraph)
        }
    }
    if ($Supress -eq $true) { return } else { return $NewParagraph }
}