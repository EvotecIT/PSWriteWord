@{
    Copyright = '(c) 2011-2019 Przemyslaw Klys. All rights reserved.'
    PrivateData = @{
        PSData = @{
            LicenseUri = 'https://github.com/EvotecIT/PSWriteWord/blob/master/License'
            Tags = 'word', 'docx', 'write', 'PSWord', 'office', 'windows', 'doc'
            ProjectUri = 'https://github.com/EvotecIT/PSWriteWord'
            IconUri = 'https://evotec.xyz/wp-content/uploads/2018/10/PSWriteWord.png'
            Prerelease = 'Preview6'
        }
    }
    Description = 'Simple project to create Microsoft Word in PowerShell without having Office installed.'
    PowerShellVersion = '5.1'
    FunctionsToExport = 'New-WordBlock', 'New-WordBlockList', 'New-WordBlockPageBreak', 'New-WordBlockParagraph', 'New-WordBlockTable', 'Add-WordBarChart', 'Add-WordChartSeries', 'Add-WordLineChart', 'Add-WordPieChart', 'Add-WordFooter', 'Add-WordHeader', 'Get-WordFooter', 'Get-WordHeader', 'Add-WordHyperLink', 'Set-WordHyperLink', 'Add-WordList', 'Add-WordListItem', 'Convert-ListToHeadings', 'New-WordList', 'New-WordListItem', 'New-WordListItemInternal', 'Set-WordList', 'Get-WordDocument', 'Merge-WordDocument', 'New-WordDocument', 'Save-WordDocument', 'Add-WordCustomProperty', 'Add-WordEquation', 'Add-WordLine', 'Add-WordPageCount', 'Add-WordProtection', 'Add-WordSection', 'Add-WordTabStopPosition', 'Get-WordCustomProperty', 'Get-WordPageSettings', 'Get-WordSection', 'Set-WordMargins', 'Set-WordOrientation', 'Set-WordPageSettings', 'Set-WordPageSize', 'Add-WordPageBreak', 'Add-WordParagraph', 'Add-WordText', 'Get-WordListItemParagraph', 'Get-WordParagraphs', 'Get-WordParagraphForList', 'Remove-WordParagraph', 'Remove-WordText', 'Set-WordParagraph', 'Set-WordText', 'Set-WordTextAlignment', 'Set-WordTextBold', 'Set-WordTextCapsStyle', 'Set-WordTextColor', 'Set-WordTextDirection', 'Set-WordTextFontFamily', 'Set-WordTextFontSize', 'Set-WordTextHeadingType', 'Set-WordTextHidden', 'Set-WordTextHighlight', 'Set-WordTextIndentationFirstLine', 'Set-WordTextIndentationHanging', 'Set-WordTextItalic', 'Set-WordTextKerning', 'Set-WordTextLanguage', 'Set-WordTextMisc', 'Set-WordTextPercentageScale', 'Set-WordTextPosition', 'Set-WordTextScript', 'Set-WordTextShadingType', 'Set-WordTextSpacing', 'Set-WordTextSpacingAfter', 'Set-WordTextSpacingBefore', 'Set-WordTextStrikeThrough', 'Set-WordTextUnderlineColor', 'Set-WordTextUnderlineStyle', 'Add-WordPicture', 'Get-WordPicture', 'Remove-WordPicture', 'Set-WordPicture', 'Add-WordTable', 'Add-WordTableCellValue', 'Add-WordTableColumn', 'Add-WordTableRow', 'Add-WordTableTitle', 'Copy-WordTableRow', 'Get-WordTable', 'Get-WordTableRow', 'New-WordTable', 'New-WordTableBorder', 'Remove-WordTable', 'Remove-WordTableColumn', 'Remove-WordTableRow', 'Set-WordTable', 'Set-WordTableAutoFit', 'Set-WordTableBorder', 'Set-WordTableCell', 'Set-WordTableCellFillColor', 'Set-WordTableCellShadingColor', 'Set-WordTableColumnWidth', 'Set-WordTableColumnWidthByIndex', 'Set-WordTableDesign', 'Set-WordTableDirection', 'Set-WordTablePageBreak', 'Set-WordTableRowMergeCells', 'Add-WordTOC', 'Add-WordTocItem'
    Author = 'Przemyslaw Klys'
    RequiredModules = @{
        ModuleVersion = '0.0.79'
        ModuleName = 'PSSharedGoods'
        GUID = 'ee272aa8-baaa-4edf-9f45-b6d6f7d844fe'
    }
    GUID = '6314c78a-d011-4489-b462-91b05ec6a5c4'
    RootModule = 'PSWriteWord.psm1'
    AliasesToExport = 'Add-WordPageNumber'
    ModuleVersion = '1.0.0'
    CompanyName = 'Evotec'
}