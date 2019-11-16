function Add-WordHyperLink {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Document.NET.Container]$WordDocument,
        [string] $UrlText,
        [string] $UrlLink,
        [bool] $Supress = $false
    )
    $Url = New-Object -TypeName Uri -ArgumentList $UrlLink

    return $WordDocument.AddHyperlink( $UrlText, $Url )
}