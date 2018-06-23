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
        [string] $FilePath,
        [string] $Language
    )

    if (-not [string]::IsNullOrEmpty($Language)) {
        Write-Verbose "Save-WordDocument - Setting Language to $Language"
        $Paragraphs = Get-WordParagraphs -WordDocument $WordDocument
        foreach ($p in $Paragraphs) {
            Set-WordParagraph -Paragraph $p -Language $Language
        }
    }

    if ([string]::IsNullOrEmpty($FilePath)) {
        $WordDocument.Save()
    } else {
        $WordDocument.SaveAs($FilePath)
    }
}