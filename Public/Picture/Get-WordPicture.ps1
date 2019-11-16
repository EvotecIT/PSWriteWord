function Get-WordPicture {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Document.NET.Container]$WordDocument,
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Document.NET.InsertBeforeOrAfter] $Paragraph,
        [switch] $ListParagraphs,
        [switch] $ListPictures,
        [nullable[int]] $PictureID
    )
    if ($ListParagraphs -eq $true -and $ListPictures -eq $true) {
        throw 'Only one option is possible at time (-ListParagraphs or -ListPictures)'
    }
    if ($ListParagraphs) {
        $Paragraphs = $WordDocument.Paragraphs
        $List = foreach ($p in $Paragraphs) {
            if ($p.Pictures -ne $null) {
                $p
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
