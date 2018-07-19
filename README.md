[![Build status](https://ci.appveyor.com/api/projects/status/5ib62bbpnj92wcd7?svg=true)](https://ci.appveyor.com/project/PrzemyslawKlys/pswriteword)[![Build status](https://img.shields.io/powershellgallery/v/PSWriteWord.svg)](https://www.powershellgallery.com/packages/PSWriteWord)[![Build status](https://img.shields.io/powershellgallery/dt/PSWriteWord.svg)](https://www.powershellgallery.com/packages/PSWriteWord)

###### PSWriteWord - PowerShell Module
*PSWriteWord* is powershell module to create Microsoft Word documents without Microsoft Word installed.

Overview of this module: https://evotec.xyz/hub/scripts/pswriteword-powershell-module/

###### Updates:
- 0.4.7 - Added -ContinueFormatting to Add-WordText - same implementation as for Add-WordTable
- 0.4.6 - https://evotec.xyz/pswriteword-add-wordtable-add-wordtext-expanded/
- 0.4.1 - https://evotec.xyz/pswriteword-updated-to-0-4-1-breaking-change-included/

###### Example usage of Add-WordTable / Add-WordText in action

![image](https://evotec.xyz/wp-content/uploads/2018/07/PSWriteWord-MoreAction.gif.pagespeed.ce.ULhhEhkC5h.gif)

###### Requirements:
- Works only on Windows (as NET CORE is not supported by DLL) - until Xceed ads that to their version. Which will take a while for the free version to have (if ever).

###### Tested on: (feel free to provide what you use it on)
- Windows 10 1803+

###### Credits
This module is based on DocX from Xceed Software (https://github.com/xceedsoftware/DocX). DocX is the free, open source version of Xceed Words for .NET. Originally written by Cathal Coffey, and maintained by Przemyslaw Klys (me), it is now maintained by Xceed. This also means any bugs / issues with C# version of it will be available in PSWriteWord. On the other hand since Xceed releases new versions of DocX (just a few versions behind the pay version) it means eventually PSWriteWord will get even more features. This also means that if something is not yet available in commands that are listed below or are incomplete you can take the .net approach and simply use it that way.

###### Dedicated word functions

```
CommandType     Name                                               Version    Source
-----------     ----                                               -------    ------
Function        Add-WordBarChart                                   0.4.6      PSWriteWord
Function        Add-WordBookmark                                   0.4.6      PSWriteWord
Function        Add-WordChartSeries                                0.4.6      PSWriteWord
Function        Add-WordCustomProperty                             0.4.6      PSWriteWord
Function        Add-WordEquation                                   0.4.6      PSWriteWord
Function        Add-WordFooter                                     0.4.6      PSWriteWord
Function        Add-WordHeader                                     0.4.6      PSWriteWord
Function        Add-WordHyperLink                                  0.4.6      PSWriteWord
Function        Add-WordImage                                      0.4.6      PSWriteWord
Function        Add-WordLine                                       0.4.6      PSWriteWord
Function        Add-WordLineChart                                  0.4.6      PSWriteWord
Function        Add-WordList                                       0.4.6      PSWriteWord
Function        Add-WordParagraph                                  0.4.6      PSWriteWord
Function        Add-WordPicture                                    0.4.6      PSWriteWord
Function        Add-WordPieChart                                   0.4.6      PSWriteWord
Function        Add-WordProtection                                 0.4.6      PSWriteWord
Function        Add-WordSection                                    0.4.6      PSWriteWord
Function        Add-WordTable                                      0.4.6      PSWriteWord
Function        Add-WordTableCellValue                             0.4.6      PSWriteWord
Function        Add-WordTableColumn                                0.4.6      PSWriteWord
Function        Add-WordTableRow                                   0.4.6      PSWriteWord
Function        Add-WordTableTitle                                 0.4.6      PSWriteWord
Function        Add-WordTabStopPosition                            0.4.6      PSWriteWord
Function        Add-WordText                                       0.4.6      PSWriteWord
Function        Add-WordTOC                                        0.4.6      PSWriteWord
Function        Copy-WordTable                                     0.4.6      PSWriteWord
Function        Copy-WordTableRow                                  0.4.6      PSWriteWord
Function        Get-WordCustomProperty                             0.4.6      PSWriteWord
Function        Get-WordDocument                                   0.4.6      PSWriteWord
Function        Get-WordFooter                                     0.4.6      PSWriteWord
Function        Get-WordHeader                                     0.4.6      PSWriteWord
Function        Get-WordPageSettings                               0.4.6      PSWriteWord
Function        Get-WordParagraphForList                           0.4.6      PSWriteWord
Function        Get-WordParagraphs                                 0.4.6      PSWriteWord
Function        Get-WordPicture                                    0.4.6      PSWriteWord
Function        Get-WordSection                                    0.4.6      PSWriteWord
Function        Get-WordTable                                      0.4.6      PSWriteWord
Function        Get-WordTableRow                                   0.4.6      PSWriteWord
Function        New-WordDocument                                   0.4.6      PSWriteWord
Function        New-WordTable                                      0.4.6      PSWriteWord
Function        New-WordTableBorder                                0.4.6      PSWriteWord
Function        Remove-WordPicture                                 0.4.6      PSWriteWord
Function        Remove-WordTable                                   0.4.6      PSWriteWord
Function        Remove-WordTableColumn                             0.4.6      PSWriteWord
Function        Remove-WordTableRow                                0.4.6      PSWriteWord
Function        Save-WordDocument                                  0.4.6      PSWriteWord
Function        Set-WordHyperLink                                  0.4.6      PSWriteWord
Function        Set-WordMargins                                    0.4.6      PSWriteWord
Function        Set-WordOrientation                                0.4.6      PSWriteWord
Function        Set-WordPageSettings                               0.4.6      PSWriteWord
Function        Set-WordPageSize                                   0.4.6      PSWriteWord
Function        Set-WordParagraph                                  0.4.6      PSWriteWord
Function        Set-WordPicture                                    0.4.6      PSWriteWord
Function        Set-WordTable                                      0.4.6      PSWriteWord
Function        Set-WordTableAutoFit                               0.4.6      PSWriteWord
Function        Set-WordTableBorder                                0.4.6      PSWriteWord
Function        Set-WordTableColumnWidth                           0.4.6      PSWriteWord
Function        Set-WordTableColumnWidthByIndex                    0.4.6      PSWriteWord
Function        Set-WordTableContinueFormatting                    0.4.6      PSWriteWord
Function        Set-WordTableDesign                                0.4.6      PSWriteWord
Function        Set-WordTableDirection                             0.4.6      PSWriteWord
Function        Set-WordTablePageBreak                             0.4.6      PSWriteWord
Function        Set-WordText                                       0.4.6      PSWriteWord
Function        Set-WordTextAlignment                              0.4.6      PSWriteWord
Function        Set-WordTextBold                                   0.4.6      PSWriteWord
Function        Set-WordTextCapsStyle                              0.4.6      PSWriteWord
Function        Set-WordTextColor                                  0.4.6      PSWriteWord
Function        Set-WordTextDirection                              0.4.6      PSWriteWord
Function        Set-WordTextFontFamily                             0.4.6      PSWriteWord
Function        Set-WordTextFontSize                               0.4.6      PSWriteWord
Function        Set-WordTextHeadingType                            0.4.6      PSWriteWord
Function        Set-WordTextHidden                                 0.4.6      PSWriteWord
Function        Set-WordTextHighlight                              0.4.6      PSWriteWord
Function        Set-WordTextIndentationFirstLine                   0.4.6      PSWriteWord
Function        Set-WordTextIndentationHanging                     0.4.6      PSWriteWord
Function        Set-WordTextItalic                                 0.4.6      PSWriteWord
Function        Set-WordTextKerning                                0.4.6      PSWriteWord
Function        Set-WordTextLanguage                               0.4.6      PSWriteWord
Function        Set-WordTextMisc                                   0.4.6      PSWriteWord
Function        Set-WordTextPercentageScale                        0.4.6      PSWriteWord
Function        Set-WordTextPosition                               0.4.6      PSWriteWord
Function        Set-WordTextScript                                 0.4.6      PSWriteWord
Function        Set-WordTextShadingType                            0.4.6      PSWriteWord
Function        Set-WordTextSpacing                                0.4.6      PSWriteWord
Function        Set-WordTextSpacingAfter                           0.4.6      PSWriteWord
Function        Set-WordTextSpacingBefore                          0.4.6      PSWriteWord
Function        Set-WordTextStrikeThrough                          0.4.6      PSWriteWord
Function        Set-WordTextUnderlineColor                         0.4.6      PSWriteWord
Function        Set-WordTextUnderlineStyle                         0.4.6      PSWriteWord
```

###### Support functions

```
CommandType     Name                                               Version    Source
-----------     ----                                               -------    ------
Function        Add-ToArray                                        0.4.6      PSWriteWord
Function        Convert-ListToHeadings                             0.4.6      PSWriteWord
Function        Convert-ObjectToProcess                            0.4.6      PSWriteWord
Function        Get-ObjectCount                                    0.4.6      PSWriteWord
Function        Get-ObjectData                                     0.4.6      PSWriteWord
Function        Get-ObjectTitles                                   0.4.6      PSWriteWord
Function        New-ArrayList                                      0.4.6      PSWriteWord
Function        Remove-FromArray                                   0.4.6      PSWriteWord

```
