function Get-WordSection {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Document.NET.Container] $WordDocument
    )
    return $WordDocument.Sections
}