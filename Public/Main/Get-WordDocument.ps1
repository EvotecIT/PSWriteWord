function Get-WordDocument {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][alias('Path')][string] $FilePath,
        [string] $LicenseKey
    )
    if ($FilePath -ne '') {
        $FilePath = Convert-Path -LiteralPath $FilePath

        if (Test-Path -LiteralPath $FilePath) {
            try {
                if ($LicenseKey) {
                    $null = [Licenser]::LicenseKey($LicenseKey)
                }
                $WordDocument = [Xceed.Words.NET.DocX]::Load($FilePath)
                Add-Member -InputObject $WordDocument -MemberType NoteProperty -Name FilePath -Value $FilePath
            } catch {
                $ErrorMessage = $_.Exception.Message
                if ($ErrorMessage -like '*Xceed.Document.NET.Licenser.LicenseKey property must be set to a valid license key in the code of your application before using this product.*') {
                    Write-Warning "Get-WordDocument - PSWriteWord on .NET CORE works only with pay version. Please provide license key."
                    return
                } else {
                    Write-Warning "Get-WordDocument - Document: $FilePath Error: $ErrorMessage"
                    return
                }
            }
        } else {
            Write-Warning "Get-WordDocument - Document doesn't exists in path $FilePath. Terminating loading word from file."
            return
        }
    }
    return $WordDocument
}