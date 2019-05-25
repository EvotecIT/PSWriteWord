function Get-WordSection {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName, ValueFromPipeline)][Container] $WordDocument
    )
    return $WordDocument.Sections
}