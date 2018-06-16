function Add-Section {
    param (
        $WordDocument,
        [switch] $PageBreak
    )
    if ($PageBreak) {
        $WordDocument.InsertSectionPageBreak()
    } else {
        $WordDocument.InsertSection()
    }
}