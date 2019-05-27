function Set-WordTableAutoFit {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][InsertBeforeOrAfter] $Table,
        [nullable[AutoFit]] $AutoFit
    )
    if (($null -ne $Table) -and ($null -ne $AutoFit)) {
        Write-Verbose "Set-WordTabelAutofit - Setting Table Autofit to: $AutoFit"
        $Table.AutoFit = $AutoFit
    }
    return $Table
}