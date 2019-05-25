function Merge-WordDocument {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][alias('Path')][string] $FilePath1,
        [alias('Append')][string] $FilePath2,
        [string] $FileOutput,
        [switch] $OpenDocument,
        [bool] $Supress = $false
    )
    if ($FilePath1 -ne '' -and $FilePath2 -ne '' -and (Test-Path -LiteralPath $FilePath1) -and (Test-Path -LiteralPath $FilePath2)) {
        try {
            $WordDocument1 = Get-WordDocument -FilePath $FilePath1
            $WordDocument2 = Get-WordDocument -FilePath $FilePath2

            $WordDocument1.InsertDocument($WordDocument2, $true)
            $FilePathOutput = Save-WordDocument -WordDocument $WordDocument1 -FilePath $FileOutput -OpenDocument:$OpenDocument
        } catch {
            $ErrorMessage = $_.Exception.Message
            if ($ErrorMessage -like '*Xceed.Document.NET.Licenser.LicenseKey property must be set to a valid license key in the code of your application before using this product.*') {
                Write-Warning "Merge-WordDocument - PSWriteWord on .NET CORE works only with pay version. Please provide license key."
                Exit
            } else {
                Write-Warning "Merge-WordDocument - Error: $ErrorMessage"
                Exit
            }
        }
        if (-not $Supress) { return $FilePathOutput }
    } else {
        Write-Warning "Merge-WordDocument - Either $FilePath1 or $FilePath2 doesn't exists. Terminating."
    }
}