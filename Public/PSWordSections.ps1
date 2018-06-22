function Add-WordSection {
    [CmdletBinding()]
    param (
        [Xceed.Words.NET.Container] $WordDocument,
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