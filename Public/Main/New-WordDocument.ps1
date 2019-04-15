function New-WordDocument {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][alias('Path')][string] $FilePath = ''
    )
    $Word = [Xceed.Words.NET.DocX]
    $WordDocument = $Word::Create($FilePath)
    $WordDocument | Add-Member -MemberType NoteProperty -Name FilePath -Value $FilePath
    return $WordDocument
}