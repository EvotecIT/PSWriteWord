function New-WordProcessing {
    [CmdletBinding()]
    param(
        [Array] $Content,
        [Xceed.Document.NET.Container] $WordDocument
    )
    if ($Content.Count -gt 0) {
        foreach ($Parameters in $Content) {
            if ($Parameters.ObjectType -eq 'List') {
                New-DocWordList -WordDocument $WordDocument -Parameters $Parameters
            } elseif ($Parameters.ObjectType -eq 'Table') {
                New-DocWordTable -WordDocument $WordDocument -Parameters $Parameters
            } elseif ($Parameters.ObjectType -eq 'TOC') {
                New-DocWordTOC -WordDocument $WordDocument -Parameters $Parameters
            } elseif ($Parameters.ObjectType -eq 'Text') {
                New-DocWordText -WordDocument $WordDocument -Parameters $Parameters
            } elseif ($Parameters.ObjectType -eq 'TocItem') {
                New-DocWordNumbering -WordDocument $WordDocument -Parameters $Parameters
            } elseif ($Parameters.ObjectType -eq 'PageBreak') {
                New-DocWordPageBreak -WordDocument $WordDocument -Parameters $Parameters
            } elseif ($Parameters.ObjectType -eq 'ChartPie') {
                New-DocWordChart -WordDocument $WordDocument -Parameters $Parameters
            }
        }
    }
}