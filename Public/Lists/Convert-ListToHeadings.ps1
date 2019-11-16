function Convert-ListToHeadings {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Document.NET.Container] $WordDocument,
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Document.NET.InsertBeforeOrAfter] $List,
        [alias ("HT")] [Xceed.Document.NET.HeadingType] $HeadingType = [Xceed.Document.NET.HeadingType]::Heading1,
        [bool] $Supress
    )
    Write-Verbose "Convert-ListToHeadings - NumID: $($List.NumID)"
    $Paragraphs = Get-WordParagraphForList -WordDocument $WordDocument -ListID $List.NumID
    Write-Verbose "Convert-ListToHeadings - List Elements Count: $($Paragraphs.Count)"
    $ParagraphsWithHeadings = foreach ($p in $Paragraphs) {
        Write-Verbose "Convert-ListToHeadings - Loop: $HeadingType"
        $p.StyleName = $HeadingType
        $p
    }
    if ($Supress) { return } else { return $ParagraphsWithHeadings }
}