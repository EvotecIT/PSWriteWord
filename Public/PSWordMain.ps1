function New-WordDocument {
    [CmdletBinding()]
    param(
        [string] $FilePath = ''
    )
    $Word = [Xceed.Words.NET.DocX]
    $WordDocument = $Word::Create($FilePath)
    return $WordDocument
}

function Get-WordDocument {
    [CmdletBinding()]
    param(
        $FilePath
    )
    $Word = [Xceed.Words.NET.DocX]
    $WordDocument = $Word::Load($FilePath)
    return $WordDocument
}

function Save-WordDocument {
    [CmdletBinding()]
    param (
        [Xceed.Words.NET.Container] $WordDocument,
        [string] $FilePath = ''
    )
    if ($FilePath -eq '') {
        $WordDocument.Save()
    } else {
        $WordDocument.SaveAs($FilePath)
    }
}