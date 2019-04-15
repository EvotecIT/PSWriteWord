function Merge-WordDocument {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][alias('Path')][string] $FilePath1,
        [alias('Append')][string] $FilePath2,
        [string] $FileOutput,
        [switch] $OpenDocument,
        [bool] $Supress = $false
    )
    $WordDocument1 = Get-WordDocument -FilePath $FilePath1
    $WordDocument2 = Get-WordDocument -FilePath $FilePath2

    $WordDocument1.InsertDocument($WordDocument2, $true)

    $FilePathOutput = Save-WordDocument -WordDocument $WordDocument1 -FilePath $FileOutput -OpenDocument:$OpenDocument
    if ($Supress) { return } else { return $FilePathOutput }
}