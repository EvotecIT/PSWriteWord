function Set-WordTextStyle {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Document.NET.InsertBeforeOrAfter] $Paragraph,
        [string] $StyleName,
        [bool] $Supress = $false
    )
    if ($null -ne $Paragraph -and $StyleName -ne $null -and $StyleName -ne '') {
        Write-Verbose "Set-WordTextStyle - Setting StyleName to $StyleName"
        $Paragraph.StyleName = $StyleName
    }
    if ($Supress) { return } else { return $Paragraph }
}