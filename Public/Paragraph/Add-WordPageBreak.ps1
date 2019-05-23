function Add-WordPageBreak {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.Container]$WordDocument,
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][alias('Paragraph', 'Table', 'List')][Xceed.Words.NET.InsertBeforeOrAfter] $WordObject,
        [alias('Insert')][Xceed.Words.NET.InsertBeforeOrAfter] $InsertWhere = [Xceed.Words.NET.InsertBeforeOrAfter]::AfterSelf,
        [bool] $Supress = $false
    )
    $RemovalRequired = $false
    if ($WordObject -eq $null) {
        Write-Verbose "Add-WordPageBreak - Adding temporary paragraph"
        $RemovalRequired = $True
        $WordObject = $WordDocument.InsertParagraph()
    }
    if ($InsertWhere -eq [Xceed.Words.NET.InsertBeforeOrAfter]::AfterSelf) {
        $WordObject.InsertPageBreakAfterSelf()
    } else {
        $WordObject.InsertPageBreakBeforeSelf()
    }
    if ($RemovalRequired) {
        Write-Verbose "Add-WordPageBreak - Removing paragraph that was added temporary"
        Remove-WordParagraph -Paragraph $WordObject
    }
    if ($Supress -eq $true) { return } else { return $WordObject }
}