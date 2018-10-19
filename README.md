# PSWriteWord - PowerShell Module

[![Build status](https://ci.appveyor.com/api/projects/status/5ib62bbpnj92wcd7?svg=true)](https://ci.appveyor.com/project/PrzemyslawKlys/pswriteword)
<!--
[![Build status](https://img.shields.io/powershellgallery/v/PSWriteWord.svg)](https://www.powershellgallery.com/packages/PSWriteWord)
[![Build status](https://img.shields.io/powershellgallery/dt/PSWriteWord.svg)](https://www.powershellgallery.com/packages/PSWriteWord)
-->

*PSWriteWord* is PowerShell module to create Microsoft Word documents without Microsoft Word installed.

Overview of this module: https://evotec.xyz/hub/scripts/pswriteword-powershell-module/

## Updates:

- 0.5.7 - 19.10.2018
    - Addition for Add-WordList to support singular (string,bool etc types)
    - Added Add-WordList Tests
- 0.5.6 - Bugfixes
- 0.5.4 - Added **NoLegend** switch to Charts (Add-WordPieChart, Add-WordLineChart, Add-WordBarChart)
- 0.5.1 - Small cleanup
- 0.5.0 - https://evotec.xyz/pswriteword-version-0-5-1/
- 0.4.7 - Added -ContinueFormatting to Add-WordText - same implementation as for Add-WordTable
- 0.4.6 - https://evotec.xyz/pswriteword-add-wordtable-add-wordtext-expanded/
- 0.4.1 - https://evotec.xyz/pswriteword-updated-to-0-4-1-breaking-change-included/

## Example usage of Add-WordTable / Add-WordText in action

![image](https://evotec.xyz/wp-content/uploads/2018/07/PSWriteWord-MoreAction.gif.pagespeed.ce.ULhhEhkC5h.gif)

## Requirements:

- Works only on Windows (as NET CORE is not supported by DLL) - until Xceed ads that to their version. Which will take a while for the free version to have (if ever).

## Tested on: (feel free to provide what you use it on)

- Windows 10 1803+

## Credits

This module is based on **DocX** from **Xceed Software** (<https://github.com/xceedsoftware/DocX>). DocX is the free, open source version of **Xceed Words for .NET**. Originally written by Cathal Coffey, and maintained by Przemyslaw Klys (me), it is now maintained by **Xceed**. This also means any bugs / issues with C# version of it will be available in **PSWriteWord**. On the other hand since Xceed releases new versions of DocX (just a few versions behind the pay version) it means eventually **PSWriteWord** will get even more features. This also means that if something is not yet available in commands that are listed below or are incomplete you can take the .net approach and simply use it that way.

## Commands

### Dedicated word functions

```powershell
CommandType     Name                                               Version    Source
-----------     ----                                               -------    ------
Function        Add-WordBarChart                                   0.5.0      PSWriteWord
Function        Add-WordBookmark                                   0.5.0      PSWriteWord
Function        Add-WordChartSeries                                0.5.0      PSWriteWord
Function        Add-WordCustomProperty                             0.5.0      PSWriteWord
Function        Add-WordEquation                                   0.5.0      PSWriteWord
Function        Add-WordFooter                                     0.5.0      PSWriteWord
Function        Add-WordHeader                                     0.5.0      PSWriteWord
Function        Add-WordHyperLink                                  0.5.0      PSWriteWord
Function        Add-WordImage                                      0.5.0      PSWriteWord
Function        Add-WordLine                                       0.5.0      PSWriteWord
Function        Add-WordLineChart                                  0.5.0      PSWriteWord
Function        Add-WordList                                       0.5.0      PSWriteWord
Function        Add-WordListItem                                   0.5.0      PSWriteWord
Function        Add-WordPageBreak                                  0.5.0      PSWriteWord
Function        Add-WordParagraph                                  0.5.0      PSWriteWord
Function        Add-WordPicture                                    0.5.0      PSWriteWord
Function        Add-WordPieChart                                   0.5.0      PSWriteWord
Function        Add-WordProtection                                 0.5.0      PSWriteWord
Function        Add-WordSection                                    0.5.0      PSWriteWord
Function        Add-WordTable                                      0.5.0      PSWriteWord
Function        Add-WordTableCellValue                             0.5.0      PSWriteWord
Function        Add-WordTableColumn                                0.5.0      PSWriteWord
Function        Add-WordTableRow                                   0.5.0      PSWriteWord
Function        Add-WordTableTitle                                 0.5.0      PSWriteWord
Function        Add-WordTabStopPosition                            0.5.0      PSWriteWord
Function        Add-WordText                                       0.5.0      PSWriteWord
Function        Add-WordTOC                                        0.5.0      PSWriteWord
Function        Add-WordTocItem                                    0.5.0      PSWriteWord
Function        Copy-WordTable                                     0.5.0      PSWriteWord
Function        Copy-WordTableRow                                  0.5.0      PSWriteWord
Function        Format-WordListItem                                0.5.0      PSWriteWord
Function        Get-WordCustomProperty                             0.5.0      PSWriteWord
Function        Get-WordDocument                                   0.5.0      PSWriteWord
Function        Get-WordFooter                                     0.5.0      PSWriteWord
Function        Get-WordHeader                                     0.5.0      PSWriteWord
Function        Get-WordListItemParagraph                          0.5.0      PSWriteWord
Function        Get-WordPageSettings                               0.5.0      PSWriteWord
Function        Get-WordParagraphForList                           0.5.0      PSWriteWord
Function        Get-WordParagraphs                                 0.5.0      PSWriteWord
Function        Get-WordPicture                                    0.5.0      PSWriteWord
Function        Get-WordSection                                    0.5.0      PSWriteWord
Function        Get-WordTable                                      0.5.0      PSWriteWord
Function        Get-WordTableRow                                   0.5.0      PSWriteWord
Function        New-WordBlock                                      0.5.0      PSWriteWord
Function        New-WordBlockList                                  0.5.0      PSWriteWord
Function        New-WordBlockPageBreak                             0.5.0      PSWriteWord
Function        New-WordBlockParagraph                             0.5.0      PSWriteWord
Function        New-WordBlockTable                                 0.5.0      PSWriteWord
Function        New-WordDocument                                   0.5.0      PSWriteWord
Function        New-WordListItem                                   0.5.0      PSWriteWord
Function        New-WordTable                                      0.5.0      PSWriteWord
Function        New-WordTableBorder                                0.5.0      PSWriteWord
Function        Remove-WordParagraph                               0.5.0      PSWriteWord
Function        Remove-WordPicture                                 0.5.0      PSWriteWord
Function        Remove-WordTable                                   0.5.0      PSWriteWord
Function        Remove-WordTableColumn                             0.5.0      PSWriteWord
Function        Remove-WordTableRow                                0.5.0      PSWriteWord
Function        Remove-WordText                                    0.5.0      PSWriteWord
Function        Save-WordDocument                                  0.5.0      PSWriteWord
Function        Set-WordContinueFormatting                         0.5.0      PSWriteWord
Function        Set-WordHyperLink                                  0.5.0      PSWriteWord
Function        Set-WordList                                       0.5.0      PSWriteWord
Function        Set-WordMargins                                    0.5.0      PSWriteWord
Function        Set-WordOrientation                                0.5.0      PSWriteWord
Function        Set-WordPageSettings                               0.5.0      PSWriteWord
Function        Set-WordPageSize                                   0.5.0      PSWriteWord
Function        Set-WordParagraph                                  0.5.0      PSWriteWord
Function        Set-WordPicture                                    0.5.0      PSWriteWord
Function        Set-WordTable                                      0.5.0      PSWriteWord
Function        Set-WordTableAutoFit                               0.5.0      PSWriteWord
Function        Set-WordTableBorder                                0.5.0      PSWriteWord
Function        Set-WordTableCell                                  0.5.0      PSWriteWord
Function        Set-WordTableCellFillColor                         0.5.0      PSWriteWord
Function        Set-WordTableCellShadingColor                      0.5.0      PSWriteWord
Function        Set-WordTableColumnWidth                           0.5.0      PSWriteWord
Function        Set-WordTableColumnWidthByIndex                    0.5.0      PSWriteWord
Function        Set-WordTableDesign                                0.5.0      PSWriteWord
Function        Set-WordTableDirection                             0.5.0      PSWriteWord
Function        Set-WordTablePageBreak                             0.5.0      PSWriteWord
Function        Set-WordTableRowMergeCells                         0.5.0      PSWriteWord
Function        Set-WordText                                       0.5.0      PSWriteWord
Function        Set-WordTextAlignment                              0.5.0      PSWriteWord
Function        Set-WordTextBold                                   0.5.0      PSWriteWord
Function        Set-WordTextCapsStyle                              0.5.0      PSWriteWord
Function        Set-WordTextColor                                  0.5.0      PSWriteWord
Function        Set-WordTextDirection                              0.5.0      PSWriteWord
Function        Set-WordTextFontFamily                             0.5.0      PSWriteWord
Function        Set-WordTextFontSize                               0.5.0      PSWriteWord
Function        Set-WordTextHeadingType                            0.5.0      PSWriteWord
Function        Set-WordTextHidden                                 0.5.0      PSWriteWord
Function        Set-WordTextHighlight                              0.5.0      PSWriteWord
Function        Set-WordTextIndentationFirstLine                   0.5.0      PSWriteWord
Function        Set-WordTextIndentationHanging                     0.5.0      PSWriteWord
Function        Set-WordTextItalic                                 0.5.0      PSWriteWord
Function        Set-WordTextKerning                                0.5.0      PSWriteWord
Function        Set-WordTextLanguage                               0.5.0      PSWriteWord
Function        Set-WordTextMisc                                   0.5.0      PSWriteWord
Function        Set-WordTextPercentageScale                        0.5.0      PSWriteWord
Function        Set-WordTextPosition                               0.5.0      PSWriteWord
Function        Set-WordTextScript                                 0.5.0      PSWriteWord
Function        Set-WordTextShadingType                            0.5.0      PSWriteWord
Function        Set-WordTextSpacing                                0.5.0      PSWriteWord
Function        Set-WordTextSpacingAfter                           0.5.0      PSWriteWord
Function        Set-WordTextSpacingBefore                          0.5.0      PSWriteWord
Function        Set-WordTextStrikeThrough                          0.5.0      PSWriteWord
Function        Set-WordTextText                                   0.5.0      PSWriteWord
Function        Set-WordTextUnderlineColor                         0.5.0      PSWriteWord
Function        Set-WordTextUnderlineStyle                         0.5.0      PSWriteWord
```

### Support functions - shouldn't be used but are there for support purposes.

```powershell
CommandType     Name                                               Version    Source
-----------     ----                                               -------    ------
Function        Add-ToArray                                        0.5.0      PSWriteWord
Function        Add-ToArrayAdvanced                                0.5.0      PSWriteWord
Function        Convert-ListToHeadings                             0.5.0      PSWriteWord
Function        Convert-ObjectToProcess                            0.5.0      PSWriteWord
Function        ConvertTo-HashtableFromPsCustomObject              0.5.0      PSWriteWord
Function        ConvertTo-PsCustomObjectFromHashtable              0.5.0      PSWriteWord
Function        Format-PSTable                                     0.5.0      PSWriteWord
Function        Format-PSTableConvertType1                         0.5.0      PSWriteWord
Function        Format-PSTableConvertType2                         0.5.0      PSWriteWord
Function        Format-PSTableConvertType3                         0.5.0      PSWriteWord
Function        Format-TransposeTable                              0.5.0      PSWriteWord
Function        Get-ColorFromARGB                                  0.5.0      PSWriteWord
Function        Get-ObjectCount                                    0.5.0      PSWriteWord
Function        Get-ObjectData                                     0.5.0      PSWriteWord
Function        Get-ObjectTitles                                   0.5.0      PSWriteWord
Function        Get-ObjectType                                     0.5.0      PSWriteWord
Function        New-ArrayList                                      0.5.0      PSWriteWord
Function        Remove-FromArray                                   0.5.0      PSWriteWord
Function        Show-Array                                         0.5.0      PSWriteWord
Function        Show-TableVisualization                            0.5.0      PSWriteWord

```
