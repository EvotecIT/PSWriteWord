function New-WordBlockList {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline, Mandatory = $true)][Xceed.Words.NET.Container]$WordDocument,
        # [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)]$Paragraph,
        [bool] $TocEnable,
        [string] $TocText,
        [int] $TocListLevel,
        [ListItemType] $TocListItemType,
        [HeadingType] $TocHeadingType,
        [int] $EmptyParagraphsBefore,
        [int] $EmptyParagraphsAfter,
        [string] $Text,
        [string] $TextListEmpty,

        [Object] $ListData,
        [ListItemType] $ListType
        # [bool] $Supress
    )
    if ($TocEnable) {
        $TOC = $WordDocument | Add-WordTocItem -Text $TocText -ListLevel $TocListLevel -ListItemType $TocListItemType -HeadingType $TocHeadingType
    }
    $WordDocument | New-WordBlockParagraph -EmptyParagraphs $EmptyParagraphsBefore
    $Paragraph = Add-WordText -WordDocument $WordDocument -Paragraph $Paragraph -Text $Text
    if ((Get-ObjectCount $ListData) -gt 0) {
        $List = Add-WordList -WordDocument $WordDocument -ListType $ListType -Paragraph $Paragraph -ListData $ListData #-Verbose
    } else {
        $Paragraph = Add-WordText -WordDocument $WordDocument -Paragraph $Paragraph -Text $TextListEmpty
    }
    $WordDocument |New-WordBlockParagraph -EmptyParagraphs $EmptyParagraphsAfter
    #if ($Supress) { return } else { return $WordDocument }
}