function New-DocWordTable {
    [CmdletBinding()]
    param(
        [Xceed.Document.NET.Container] $WordDocument,
        [PSCustomObject] $Parameters
    )

    if ($Parameters.OverWriteTitle) {
        [Xceed.Document.NET.Alignment] $TitleAlignment = $Parameters.OverwriteTitleAlignment
        [nullable[System.Drawing.KnownColor]] $TitleColor = $Parameters.OverwriteTitleColor

        $Table = Add-WordTable -WordDocument $WordDocument -Supress $false -DataTable $Parameters.DataTable -Design $Parameters.Design -AutoFit $Parameters.AutoFit -MaximumColumns $Parameters.MaximumColumns
        $Table = Set-WordTableRowMergeCells -Table $Table -RowNr 0 -MergeAll  # -ColumnNrStart 0 -ColumnNrEnd 1
        $TableParagraph = Get-WordTableRow -Table $Table -RowNr 0 -ColumnNr 0
        $TableParagraph = Set-WordText -Paragraph $TableParagraph -Text $Parameters.OverwriteTitle -Alignment $TitleAlignment -Color $TitleColor
    } else {
        $Table = Add-WordTable -WordDocument $WordDocument -Supress $true -DataTable $Parameters.DataTable -Design $Parameters.Design -AutoFit $Parameters.AutoFit -Transpose:$Parameters.Transpose
    }
}