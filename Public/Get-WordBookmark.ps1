function Get-WordBookmark {
    [cmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Document.NET.Container] $WordDocument
    )
    $WordDocument.Bookmarks | Select-Object -Property Name, Paragraph
}