function Set-WordBookmark {
    [cmdletBinding()]
    param(
        [string] $BookmarkName,
        [string] $BookmarkText
    )
    $WordDocument.InsertAtBookmark($BookmarkText, $BookmarkName)
}