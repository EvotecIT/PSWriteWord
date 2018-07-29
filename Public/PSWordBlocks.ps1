function New-WordBlock {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)]$WordDocument,
        # [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)]$Paragraph,
        [bool] $TocEnable,
        [string] $TocText,
        [int] $TocListLevel,
        [ListItemType] $TocListItemType,
        [HeadingType] $TocHeadingType,
        [int] $EmptyParagraphsBefore,
        [int] $EmptyParagraphsAfter,
        [string] $Text
    )
    if ($TocEnable) {
        $TOC = $WordDocument | Add-WordTocItem -Text $TocText -ListLevel $TocListLevel -ListItemType $TocListItemType -HeadingType $TocHeadingType
    }
    $WordDocument | New-WordBlockParagraph -EmptyParagraphs $EmptyParagraphsBefore
    $Paragraph = Add-WordText -WordDocument $WordDocument -Paragraph $Paragraph -Text $Text -Alignment both
    $WordDocument | New-WordBlockParagraph -EmptyParagraphs $EmptyParagraphsAfter
}
function New-WordBlockTable {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)]$WordDocument,
        # [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)]$Paragraph,
        [bool] $TocEnable,
        [string] $TocText,
        [int] $TocListLevel,
        [ListItemType] $TocListItemType,
        [HeadingType] $TocHeadingType,

        [int] $EmptyParagraphsBefore,
        [int] $EmptyParagraphsAfter,
        [string] $Text,

        [Object] $TableData,
        [TableDesign] $TableDesign,
        [bool] $TableTitleMerge = $false,
        [string] $TableTitleText,
        [Alignment] $TableTitleAlignment = 'center',
        [System.Drawing.Color] $TableTitleColor = 'Black',
        [bool] $ChartEnable,
        [string] $ChartTitle,
        $ChartKeys,
        $ChartValues,
        [ChartLegendPosition] $ChartLegendPosition = [ChartLegendPosition]::Bottom,
        [bool] $ChartLegendOverlay

    )
    if ($TocEnable) {
        $TOC = $WordDocument | Add-WordTocItem -Text $TocText -ListLevel $TocListLevel -ListItemType $TocListItemType -HeadingType $TocHeadingType
    }
    $WordDocument | New-WordBlockParagraph -EmptyParagraphs $EmptyParagraphsBefore
    $Paragraph = Add-WordText -WordDocument $WordDocument -Paragraph $Paragraph -Text $Text
    $Table = Add-WordTable -WordDocument $WordDocument -Paragraph $Paragraph -DataTable $TableData -AutoFit Window -Design $TableDesign -DoNotAddTitle:$TableTitleMerge

    if ($TableTitleMerge) {
        $Table = Set-WordTableRowMergeCells -Table $Table -RowNr 0 -ColumnNrStart 0 -ColumnNrEnd 1
        if ($TableTitleText -ne $null) {
            $TableParagraph = Get-WordTableRow -Table $Table -RowNr 0 -ColumnNr 0
            $TableParagraph = Add-WordText -WordDocument $WordDocument -Paragraph $TableParagraph -Text $TableTitleText -Alignment $TableTitleAlignment -Color $TableTitleColor -AppendToExistingParagraph
        }
    }
    if ($ChartEnable) {
        $WordDocument | New-WordBlockParagraph -EmptyParagraphs 1
        Add-WordPieChart -WordDocument $WordDocument -ChartName $ChartTitle -Names $ChartKeys -Values $ChartValues -ChartLegendPosition $ChartLegendPosition -ChartLegendOverlay $ChartLegendOverlay
    }
    $WordDocument | New-WordBlockParagraph -EmptyParagraphs $EmptyParagraphsAfter
}
function New-WordBlockList {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)]$WordDocument,
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
}
function New-WordBlockParagraph {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)]$WordDocument,
        [int] $EmptyParagraphs
    )
    $i = 0
    While ($i -lt $EmptyParagraphs) {
        Write-Verbose "New-WordBlockList - EmptyParagraphs $i"
        $Paragraph = Add-WordParagraph -WordDocument $WordDocument
        $i++
    }
}