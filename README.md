# PSWriteWord - PowerShell Module

<p align="center">

[![PowerShellGallery Version](https://img.shields.io/powershellgallery/v/PSWriteWord.svg)](https://www.powershellgallery.com/packages/PSWriteWord)
[![Build status](https://ci.appveyor.com/api/projects/status/5ib62bbpnj92wcd7?svg=true)](https://ci.appveyor.com/project/PrzemyslawKlys/pswriteword)
[![PowerShellGallery Platform](https://img.shields.io/powershellgallery/p/PSWriteWord.svg)](https://www.powershellgallery.com/packages/PSWriteWord)
[![PowerShellGallery Preview Version](https://img.shields.io/powershellgallery/vpre/PSWriteWord.svg?label=powershell%20gallery%20preview&colorB=yellow)](https://www.powershellgallery.com/packages/PSWriteWord)

</p>
<p align="center">

![Top Language](https://img.shields.io/github/languages/top/evotecit/PSWriteWord.svg)
![Code](https://img.shields.io/github/languages/code-size/evotecit/PSWriteWord.svg)
[![PowerShellGallery Downloads](https://img.shields.io/powershellgallery/dt/PSWriteWord.svg)](https://www.powershellgallery.com/packages/PSWriteWord)

</p>

<!--
[![Build status](https://img.shields.io/powershellgallery/v/PSWriteWord.svg)](https://www.powershellgallery.com/packages/PSWriteWord)
[![Build status](https://img.shields.io/powershellgallery/dt/PSWriteWord.svg)](https://www.powershellgallery.com/packages/PSWriteWord)
-->

_PSWriteWord_ is PowerShell module to create Microsoft Word documents without Microsoft Word installed.

Overview of this module: https://evotec.xyz/hub/scripts/pswriteword-powershell-module/

## Updates:

-   1.0.1 - 28.06.2019
    -   Fix for Supress in Add-WordPicture
-   1.0.0 (**Breaking Changes**) - 23.06.2019
    -   Removed custom enums, using Xceed ones instead
    -   Changed how Add-WordList works
    -   Changed how New-WordList / New-WordListItem works - be sure to review new, easier way
    -   Rewrote/fixed couple of functions
-   0.7.1 - 29.04.2019
    - Fixed Saving issue when file was in use and path was having dots in temporary path
-   0.7.0 - 17.04.2019
    - Some performance improvements
    - Includes new DocX DLL version 1.3 which added/fixed following:
        - In DocX, if the core properties part is missing in the document, it will now be created automatically instead of throwing an exception.
        - In Container, the Paragraphs properties will no longer include the fallback elements.
        - In Formatting, half point font sizes are now supported.
        - In Paragraph, the LineSpacing property will now correctly update the spacing between lines of text.
        - In Paragraph, calling the InsertText or RemoveText methods will now update the active runs.
        - In Paragraph, when an Auto spacing is used, the LineSpacingAfter and/or LineSpacingBefore properties will now return 0.
        - In Paragraph, a border can now be added to a simple text.
        - In Paragraph, a default line spacing or indentation defined on the document will now be used when the paragraph doesn’t specify those values.
        - In Paragraph, adding a bookmark with the same name as a previously deleted bookmark will no longer cause an exception.
        - In Picture, its size will now be defined by using the default 96 pixels per inch.
        - In Picture, the new HeightInches and WidthInches properties can now be used to get/set the image size in inches.
        - In Table, Cell will now default to a Top vertical alignment.

-   0.6.0 - 13.01.2019
    -   Fixed merging of columns/cells in Table in circumstances with multiple paragraphs in a cell
    -   Published with merging all files into single PSM1 and optimized PSD1 file which speeds up greatly Import-Module process
        If you want to read more about why I did it: https://evotec.xyz/powershell-single-psm1-file-versus-multi-file-modules/

-   0.5.17 - 13.01.2019
    -   Fixed an empty chart problem in Word Blocks. If values/keys were empty it would create an empty chart preventing Word saving feature to work.
    -   Fixed merging of columns/cells in Table
-   0.5.16 - 9.01.2019
    -   Fixed Get-WordDocument support for path with special characters
-   0.5.15 - 29.12.2018
    -   Added Add-WordPageCount with alias Add-WordPageNumber
    -   Added some examples and tests for above
-   0.5.14 - 28.12.2018
    -   Expanded support for Headers/Footers
    -   Expanded Add-WordText with -Headers/Footers
    -   Added some examples and tests
-   0.5.13 - 8.12.2018
    -   Fix Set-WordTableRowMergeCells
    -   Added Set-WordTableRowMergeCells TextMerge with Separator allowing for merging text from merged columns
    -   Fix for Remove-WordText
    -   Added tests for Set-WordTableRowMergeCells - does some additional testing to Tables just in case
-   0.5.12 - 8.12.2018
    -   Fix for Set-WordTableRowMergeCells
-   0.5.11 - 30.11.2018
    -   Added Alignment to Add-WordPicture
-   0.5.10 - 29.11.2018
    -   Added Merge-WordDocument - which brings merging of Word Documents functionality
-   0.5.9 - 5.11.2018
    -   Fix for Add-WordPicture
-   0.5.8 - 19.10.2018
    -   Added small feature New-WordBlock for PSWinDocumentation
-   0.5.7 - 19.10.2018
    -   Addition for Add-WordList to support singular (string,bool etc types)
    -   Added Add-WordList Tests
-   0.5.6 - Bugfixes
-   0.5.4 - Added **NoLegend** switch to Charts (Add-WordPieChart, Add-WordLineChart, Add-WordBarChart)
-   0.5.1 - Small cleanup
-   0.5.0 - https://evotec.xyz/pswriteword-version-0-5-1/
-   0.4.7 - Added -ContinueFormatting to Add-WordText - same implementation as for Add-WordTable
-   0.4.6 - https://evotec.xyz/pswriteword-add-wordtable-add-wordtext-expanded/
-   0.4.1 - https://evotec.xyz/pswriteword-updated-to-0-4-1-breaking-change-included/

## Example usage of Add-WordTable / Add-WordText in action

![image](https://evotec.xyz/wp-content/uploads/2018/07/PSWriteWord-MoreAction.gif.pagespeed.ce.ULhhEhkC5h.gif)

## Requirements:

-   Works only on Windows (as NET CORE is not supported by DLL) - until Xceed ads that to their version. Which will take a while for the free version to have (if ever).
-   Works only on PowerShell 5.1 (not tested on PowerShell 6.0 with Windows compatibility pack)

## Systems it should run on (marked those confirmed to run)

### Windows Client Systems

-   [ ] Windows 7 with Service Pack 1 - requires WMF 5.1 https://www.microsoft.com/en-us/download/details.aspx?id=54616
-   [ ] Windows 8.1
-   [x] Windows 10 1809
-   [x] Windows 10 1803

### Windows Server Systems

-   [ ] Windows 2008 R2 with Service Pack 1 - requires WMF 5.1 https://www.microsoft.com/en-us/download/details.aspx?id=54616
-   [x] Windows 2012 R2
-   [x] Windows 2016

## Credits

This module is based on **DocX** from **Xceed Software** (<https://github.com/xceedsoftware/DocX>). DocX is the free, open source version of **Xceed Words for .NET**. Originally written by Cathal Coffey, and maintained by Przemyslaw Klys (me), it is now maintained by **Xceed**. This also means any bugs / issues with C# version of it will be available in **PSWriteWord**. On the other hand since Xceed releases new versions of DocX (just a few versions behind the pay version) it means eventually **PSWriteWord** will get even more features. This also means that if something is not yet available in commands that are listed below or are incomplete you can take the .net approach and simply use it that way.

### License
**MIT License** applies only to PowerShell code/code within this repo. DLL is property of Xceed and as such is licensed under **Microsoft Public License (Ms-PL)**.

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