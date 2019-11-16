Function Set-WordParagraph {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Document.NET.InsertBeforeOrAfter] $Paragraph,
        [Xceed.Document.NET.Alignment] $Alignment,
        [Xceed.Document.NET.Direction] $Direction,
        [string] $Language,
        [bool] $Supress = $false
    )
    if ($Paragraph -ne $null) {
        #Write-Verbose "Set-WordParagraph - Paragraph is not null"
        if ($Alignment -ne $null) {
            Write-Verbose "Set-WordParagraph - Setting Alignment to $Alignment"
            $Paragraph.Alignment = $Alignment
        }
        if ($Direction -ne $null) {
            Write-Verbose "Set-WordParagraph - Setting Direction to $Direction"
            $Paragraph.Direction = $Direction
        }
        if ($Language -ne $null) {
            $Culture = [System.Globalization.CultureInfo]::GetCultureInfo($Language)
            $Paragraph = $Paragraph.Culture($Culture)
        }
    }
    if ($Supress) { return } else { return $Paragraph }
}