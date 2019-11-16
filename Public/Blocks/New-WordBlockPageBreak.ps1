function New-WordBlockPageBreak {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline, Mandatory = $true)][Xceed.Document.NET.Container]$WordDocument,
        [int] $PageBreaks,
        [bool] $Supress
    )
    $i = 0
    While ($i -lt $PageBreaks) {
        Write-Verbose "New-WordBlockPageBreak - PageBreak $i"
        $WordDocument | Add-WordPageBreak -Supress $True
        $i++
    }
}