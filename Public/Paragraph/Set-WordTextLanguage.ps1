function Set-WordTextLanguage {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][InsertBeforeOrAfter] $Paragraph,
        [string]$Language,
        [bool] $Supress = $false
    )
    if ($null -ne $Paragraph -and $Language -ne $null -and $Language -ne '') {
        $Culture = [System.Globalization.CultureInfo]::GetCultureInfo($Language)
        $Paragraph = $Paragraph.Culture($Culture)
    }
    if ($Supress) { return } else { return $Paragraph }
}