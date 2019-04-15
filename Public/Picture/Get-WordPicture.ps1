function Get-WordPicture {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.Container]$WordDocument,
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [switch] $ListParagraphs,
        [switch] $ListPictures,
        [nullable[int]] $PictureID
    )
    if ($ListParagraphs -eq $true -and $ListPictures -eq $true) {
        throw 'Only one option is possible at time (-ListParagraphs or -ListPictures)'
    }
    if ($ListParagraphs) {
        $List = New-ArrayList
        $Paragraphs = $WordDocument.Paragraphs
        foreach ($p in $Paragraphs) {
            if ($p.Pictures -ne $null) {
                Add-ToArray -List $List -Element $p
            }
        }
        return $List
    }
    if ($ListPictures) {
        return $WordDocument.Pictures
    }
    if ($PictureID -ne $null) {
        return $WordDocument.Pictures[$PictureID]
    }
}