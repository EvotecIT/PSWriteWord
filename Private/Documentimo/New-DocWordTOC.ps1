function New-DocWordTOC {
    [CmdletBinding()]
    param(
        [Xceed.Document.NET.Container] $WordDocument,
        [PSCustomObject] $Parameters
    )

    Add-WordToc -WordDocument $WordDocument -Title $Parameters.Title -Switches $Parameters.Switches -RightTabPos $Parameters.RightTabPos -Supress $True
}