function Add-WordLine {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Document.NET.Container]$WordDocument,
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Document.NET.InsertBeforeOrAfter] $Paragraph,
        [Xceed.Document.NET.HorizontalBorderPosition] $HorizontalBorderPosition = [Xceed.Document.NET.HorizontalBorderPosition]::Bottom,
        [ValidateSet('single', 'double', 'triple')] $LineType = 'single',
        [nullable[int]] $LineSize = 6,
        [nullable[int]] $LineSpace = 1,
        [string] $LineColor = 'black',
        [bool] $Supress
    )
    if ($Paragraph -eq $null) {
        $Paragraph = Add-WordParagraph -WordDocument $WordDocument -Supress $False
    }
    $Paragraph = $Paragraph.InsertHorizontalLine($HorizontalBorderPosition, $LineType, $LineSize, $LineSpace, $LineColor );
    if ($Supress) { return } else { $Paragraph }
}