function Add-WordSection {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName, ValueFromPipeline)][Container] $WordDocument,
        [switch] $PageBreak,
        [bool] $Supress
    )
    if ($PageBreak) {
        $Data = $WordDocument.InsertSectionPageBreak()
    } else {
        $Data = $WordDocument.InsertSection()
    }
    if ($Supress -eq $true) { return } else {return $Data}
}