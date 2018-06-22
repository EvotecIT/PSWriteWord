function Add-WordSection {
    [CmdletBinding()]
    param (
        [Xceed.Words.NET.Container] $WordDocument,
        [switch] $PageBreak
    )
    if ($PageBreak) {
        $WordDocument.InsertSectionPageBreak()
    } else {
        $WordDocument.InsertSection()
    }
}