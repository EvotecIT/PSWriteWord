function New-WordBlockTable {
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
        [int] $PageBreaksBefore,
        [int] $PageBreaksAfter,
        [string] $Text,

        [Object] $TableData,
        [nullable[Xceed.Words.NET.TableDesign]] $TableDesign,
        [int] $TableMaximumColumns = 5,
        [nullable[bool]] $TableTitleMerge,
        [string] $TableTitleText,
        [nullable[Xceed.Words.NET.Alignment]] $TableTitleAlignment = 'center',
        [nullable[System.Drawing.Color]] $TableTitleColor = 'Black',
        [switch] $TableTranspose,
        [nullable[bool]] $ChartEnable,
        [string] $ChartTitle,
        $ChartKeys,
        $ChartValues,
        [Xceed.Words.NET.ChartLegendPosition] $ChartLegendPosition = [Xceed.Words.NET.ChartLegendPosition]::Bottom,
        [bool] $ChartLegendOverlay
        # [bool] $Supress

    )
    $WordDocument | New-WordBlockPageBreak -PageBreaks $PageBreaksBefore
    if ($TocEnable) {
        $TOC = $WordDocument | Add-WordTocItem -Text $TocText -ListLevel $TocListLevel -ListItemType $TocListItemType -HeadingType $TocHeadingType
    }
    $WordDocument | New-WordBlockParagraph -EmptyParagraphs $EmptyParagraphsBefore
    $Paragraph = Add-WordText -WordDocument $WordDocument -Paragraph $Paragraph -Text $Text

    if ($TableData) {
        $Table = Add-WordTable -WordDocument $WordDocument -Paragraph $Paragraph -DataTable $TableData -AutoFit Window -Design $TableDesign -DoNotAddTitle:$TableTitleMerge -MaximumColumns $TableMaximumColumns -Transpose:$TableTranspose
        if ($TableTitleMerge) {
            $Table = Set-WordTableRowMergeCells -Table $Table -RowNr 0 -MergeAll  # -ColumnNrStart 0 -ColumnNrEnd 1
            if ($TableTitleText -ne $null) {
                $TableParagraph = Get-WordTableRow -Table $Table -RowNr 0 -ColumnNr 0
                $TableParagraph = Set-WordText -Paragraph $TableParagraph -Text $TableTitleText -Alignment $TableTitleAlignment -Color $TableTitleColor
            }
        }
    }
    if ($ChartEnable) {
        $WordDocument | New-WordBlockParagraph -EmptyParagraphs 1
        Add-WordPieChart -WordDocument $WordDocument -ChartName $ChartTitle -Names $ChartKeys -Values $ChartValues -ChartLegendPosition $ChartLegendPosition -ChartLegendOverlay $ChartLegendOverlay
    }
    $WordDocument | New-WordBlockParagraph -EmptyParagraphs $EmptyParagraphsAfter
    $WordDocument | New-WordBlockPageBreak -PageBreaks $PageBreaksAfter
    #if ($Supress) { return } else { return $WordDocument }
}