function Add-WordTabStopPosition {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Container]$WordDocument,
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][InsertBeforeOrAfter] $Paragraph,
        [single] $HorizontalPosition,
        [TabStopPositionLeader] $TabStopPositionLeader,
        [Alignment] $Alignment,
        [bool] $Supress = $false
    )
    if ($Paragraph -eq $null) {
        $Paragraph = Add-WordParagraph -WordDocument $WordDocument -Supress $False
    }
    $data = $Paragraph.InsertTabStopPosition($Alignment, $HorizontalPosition, $TabStopPositionLeader)

    if ($Supress) { return } else { $data }
}