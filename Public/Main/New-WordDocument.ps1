using namespace Xceed.Words.NET
using namespace Xceed.Document.NET

function New-WordDocument {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][alias('Path')][string] $FilePath = '',
        [string] $LicenseKey
    )
    try {
        if ($LicenseKey) {
            $null = [Licenser]::LicenseKey = $LicenseKey
        }
        $WordDocument = [Xceed.Words.NET.DocX]::Create($FilePath)
        Add-Member -InputObject $WordDocument -MemberType NoteProperty -Name FilePath -Value $FilePath
    } catch {
        $ErrorMessage = $_.Exception.Message
        if ($ErrorMessage -like '*Xceed.Document.NET.Licenser.LicenseKey property must be set to a valid license key in the code of your application before using this product.*') {
            Write-Warning "New-WordDocument - PSWriteWord on .NET CORE works only with pay version. Please provide license key."
            Exit
        } else {
            Write-Warning "New-WordDocument - Document: $FilePath Error: $ErrorMessage"
            Exit
        }
    }
    return $WordDocument
}