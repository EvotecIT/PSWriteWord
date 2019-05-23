function Set-WordTableAutoFit {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Table,
        [nullable[Xceed.Words.NET.AutoFit]] $AutoFit
    )
    if ($Table -ne $null -and $AutoFit -ne $null) {
        Write-Verbose "Set-WordTabelAutofit - Setting Table Autofit to: $AutoFit"
        $Table.AutoFit = $AutoFit
    }
    return $Table
}