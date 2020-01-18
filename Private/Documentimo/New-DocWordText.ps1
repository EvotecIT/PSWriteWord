function New-DocWordText {
    [CmdletBinding()]
    param(
        [Xceed.Document.NET.Container] $WordDocument,
        [PSCustomObject] $Parameters
    )


    if ($Parameters.Text) {
        Add-WordText -WordDocument $WordDocument -Text $Parameters.Text -Color $Parameters.Color -Supress $true
    }
    if ($Parameters.LineBreak) {
        Add-WordParagraph -WordDocument $WordDocument -Supress $True
    }
}