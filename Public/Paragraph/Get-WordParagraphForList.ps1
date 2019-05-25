function Get-WordParagraphForList {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Container] $WordDocument,
        [int] $ListID
    )
    $IDs = @()
    foreach ($p in $WordDocument.Paragraphs) {
        if ($p.ParagraphNumberProperties -ne $null) {
            $ListNumber = $p.ParagraphNumberProperties.LastNode.LastAttribute.Value
            if ($ListNumber -eq $ListID) {
                $IDs += $p
            }
        }
    }
    return $Ids
}