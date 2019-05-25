Function Add-WordParagraph {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Container]$WordDocument,
        [alias('Paragraph', 'Table', 'List')][InsertBeforeOrAfter] $WordObject,
        [alias('Insert')][validateset('BeforeSelf', 'AfterSelf')][string] $InsertWhere = 'AfterSelf',
        #[bool] $TrackChanges,
        [bool] $Supress = $false
    )
    $NewParagraph = $WordDocument.InsertParagraph()
    if ($WordObject -ne $null) {
        if ($InsertWhere -eq 'AfterSelf') {
            $NewParagraph = $WordObject.InsertParagraphAfterSelf($NewParagraph)
        } elseif ($InsertWhere -eq 'BeforeSelf') {
            $NewParagraph = $WordObject.InsertParagraphBeforeSelf($NewParagraph)
        }
    }
    if ($Supress -eq $true) { return } else { return $NewParagraph }
}