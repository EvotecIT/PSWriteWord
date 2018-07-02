function Add-WordTabStopPosition {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.Container]$WordDocument,
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [single] $HorizontalPosition,
        [TabStopPositionLeader] $TabStopPositionLeader,
        [Alignment] $Alignment,
        [bool] $Supress = $true
    )
    #<summary>
    #Add a new TabStopPosition in the current paragraph.
    #</summary>
    #param name="alignment">Specifies the alignment of the Tab stop.</param>
    #param name="position">Specifies the horizontal position of the tab stop.</param>
    #param name="leader">Specifies the character used to fill in the space created by a tab.</param>
    #<returns>The modified Paragraph.</returns>

    if ($Paragraph -eq $null) {
        $Paragraph = Add-WordParagraph -WordDocument $WordDocument -Supress $False
    }
    $Paragraph.InsertTabStopPosition($Alignment, $HorizontalPosition, $TabStopPositionLeader)

}