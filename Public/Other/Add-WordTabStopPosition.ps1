function Add-WordTabStopPosition {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.Container]$WordDocument,
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [single] $HorizontalPosition,
        [Xceed.Words.NET.TabStopPositionLeader] $TabStopPositionLeader,
        [Xceed.Words.NET.Alignment] $Alignment,
        [bool] $Supress = $false
    )
    if ($Paragraph -eq $null) {
        $Paragraph = Add-WordParagraph -WordDocument $WordDocument -Supress $False
    }
    $data = $Paragraph.InsertTabStopPosition($Alignment, $HorizontalPosition, $TabStopPositionLeader)

    if ($Supress) { return } else { $data }
}