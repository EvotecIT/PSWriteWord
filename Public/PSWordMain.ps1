function New-WordDocument ($FilePath = "") {
    $Word = [Xceed.Words.NET.DocX]
    $WordDocument = $Word::Create($FilePath)
    return $WordDocument
}

function Get-WordDocument($FilePath) {
    $Word = [Xceed.Words.NET.DocX]
    $WordDocument = $Word::Load($FilePath)
    return $WordDocument
}

function Save-WordDocument ($WordDocument, $FilePath = "") {
    if ($FilePath -eq "") {
        $WordDocument.Save()
    } else {
        $WordDocument.SaveAs($FilePath)
    }
    # return $WordDocument
}