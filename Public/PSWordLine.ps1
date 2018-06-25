function Add-WordLine {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.Container]$WordDocument,
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [LineType] $LineType = [LineType]::Single,
        [nullable[int]] $LineSize = 6,
        [nullable[int]] $LineSpace = 1,
        [string] $LineColor = 'black'
    )
    if ($Paragraph -eq $null) {
        $Paragraph = Add-WordParagraph -WordDocument $WordDocument -Supress $False
    }
    $Paragraph.InsertHorizontalLine($LineType, $LineSize, $LineSpace, $LineColor );
}