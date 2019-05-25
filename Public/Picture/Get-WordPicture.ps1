function Get-WordPicture {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Container]$WordDocument,
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][InsertBeforeOrAfter] $Paragraph,
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