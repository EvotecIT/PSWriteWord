function Get-WordDocument {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][alias('Path')][string] $FilePath
    )
    $Word = [Xceed.Words.NET.DocX]
    if ($FilePath -ne '') {
        if (Test-Path -LiteralPath $FilePath) {
            try {
                $WordDocument = $Word::Load($FilePath)
                $WordDocument | Add-Member -MemberType NoteProperty -Name FilePath -Value $FilePath
            } catch {
                $ErrorMessage = $_.Exception.Message
                Write-Warning "Get-WordDocument - Document: $FilePath Error: $ErrorMessage"
            }
        } else {
            Write-Warning "Get-WordDocument - Document doesn't exists in path $FilePath. Terminating loading word from file."
            return
        }
    }
    return $WordDocument
}