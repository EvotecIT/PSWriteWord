function Add-WordSection {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.Container] $WordDocument,
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

function Get-WordSection {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.Container] $WordDocument
    )
    return $WordDocument.Sections
}