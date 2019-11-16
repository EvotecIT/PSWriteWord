function Add-WordTabStopPosition {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Document.NET.Container]$WordDocument,
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Document.NET.InsertBeforeOrAfter] $Paragraph,
        [single] $HorizontalPosition,
        [Xceed.Document.NET.TabStopPositionLeader] $TabStopPositionLeader,
        [Xceed.Document.NET.Alignment] $Alignment,
        [bool] $Supress = $false
    )
    if ($null -eq $Paragraph) {
        $Paragraph = Add-WordParagraph -WordDocument $WordDocument -Supress $False
    }
    $data = $Paragraph.InsertTabStopPosition($Alignment, $HorizontalPosition, $TabStopPositionLeader)

    if ($Supress) { return } else { $data }
}