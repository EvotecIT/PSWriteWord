﻿### PSWriteWord Release History

#### - 1.1.15
  - Added `Set-WordTextStyle` - [#73](https://github.com/EvotecIT/PSWriteWord/pull/73) - tnx ylepine
  - Added `Style` parameter for `Add-WordText` to be able to apply a existing Style other than Heading 1-9 - [#73](https://github.com/EvotecIT/PSWriteWord/pull/73) - tnx ylepine
  - Added `Style` parameter for `Set-WordText` to be able to apply a existing Style other than Heading 1-9 - [#73](https://github.com/EvotecIT/PSWriteWord/pull/73) - tnx ylepine

#### - 1.1.14 - 2022.01.11
  - Removed `BreakAcrossPages` in `Add-WordTable`, `Set-WordTable`, `Set-WordTablePageBreak` as it doesn't work anymore
#### - 1.1.13 - 2021.11.28
  - Fixed bug with publishing (republished)
#### - 1.1.12 - 2021.11.20
  - 🐛 Fixed Cannot find an overload for "ReplaceText" on `Set-WordTextReplace` cmdlet [#69](https://github.com/EvotecIT/PSWriteWord/issues/69)
#### - 1.1.11 - 2021.05.16
  - 🐛 Fixed `Get-WordDocument` when working with relative paths [#53](https://github.com/EvotecIT/PSWriteWord/issues/53)
#### - 1.1.10 - 2021.05.15
  - 🐛Replaced `Exit` with `Return` in `Get-WordDocument`, and few other cmdlets to prevent PowerShell Exit
#### - 1.1.9 - 2020.09.21
  - ☑ Fixes `Set-WordTextReplace` that would not work at all
  - ☑ Updates PSD1
  - ☑ Removed dependency on PSSharedGoods when used from powershellgallery
  - ☑ Module is now signed
#### - 1.1.8 - 3.06.2020
  - Fixes `Transpose` parameter for `Add-WordTable` giving wrong order [#44](https://github.com/EvotecIT/PSWriteWord/issues/44)
#### - 1.1.7 - 6.04.2020
  - Fixes for `Add-WordCustomProperty` [#39](https://github.com/EvotecIT/PSWriteWord/issues/39)
#### - 1.1.6 - 5.04.2020
  - Fixed problems with Documentimo Colors
  - Added `Transpose` to `DocumentimoTable`
#### - 1.1.5 - 21.03.2020
  - Added Get-WordBookmark
  - Added Set-WordBookmark
#### - 1.1.4 - 8.02.2020
  - Fixes for New-WordList
#### - 1.1.3 - 18.01.2020
  - Fixes for Colors for Documentimo
#### - 1.1.2 - 18.01.2020
  - Migrated Documentimo to PSWriteWord. You can use syntax of Documentimo now.
#### - 1.1.1 - 17.11.2019
  - Fixes colors adding ability to choose them from a list (change from System.Drawing.Colors to System.Drawing.KnownColors)
  - Removes Set-WordHyperlink as it was a bit too complicated to use
  - Expanded Add-WordHyperLink with a lot of options
  - Fixed adding more than 1 hyperlink (#32)
#### - 1.1.0 - 17.11.2019
  - Removes using namespaces because of wonky way it worked
    - In case you want to keep using shortnames you would need to use both commands right after Import-Module PSWriteWord
      - using namespace Xceed.Words.NET
      - using namespace Xceed.Document.NET
  - Includes new DocX DLL version 1.4.1 which added/fixed following:
    - ☑ In Document, the new PageBackground property can now be set to customize the background color of all the document’s pages.
    - ☑ In Document, the new PageBorders property can now be set to display up to 4 different borders in a document page.
    - ☑ In Document, inserting a chart by setting its width and height is now supported.
    - ☑ In Document, adding an image from a stream will no longer throw an exception when the stream is not positioned at the beginning.
    - ☑ In Document, accessing Bookmarks multiple times no longer decreases the application’s performance.
    - ☑ In Paragraph, the MagicText property will no longer read the Xml at each call. Instead, the saved data will be used to improve the performance.
    - ☑ In Paragraph, setting an alignment and then calling InsertPicture() with no index will no longer reset the alignment to left.
    - ☑ In Paragraph, the LineSpacingBefore will no longer be added to the preceding paragraph’s LineSpacingAfter, when those values are identical.
    - ☑ In Paragraph, the AppendPageNumber and AppendPageCount methods will now return the paragraph, in order to let users continue adding actions on the paragraph.
    - ☑ In Paragraph, replacing a text without specifying a formatting will now replace all occurrences of that text.
    - ☑ In Table, using CustomTableDesignName now correctly adjusts the table’s custom style.
    - ☑ In Table, modifying the TableLook, or any property of TableLook, will now correctly update the table.

#### - 1.0.2 - 16.09.2019
  - Fix for Add-WordPicture (try/catch missing)
  - Fix for Add-WordPicture - supress would supress paragraph

#### - 1.0.1 - 28.06.2019
  - Fix for Supress in Add-WordPicture

#### - 1.0.0 (**Breaking Changes**) - 23.06.2019
  - Removed custom enums, using Xceed ones instead
  - Changed how Add-WordList works
  - Changed how New-WordList / New-WordListItem works - be sure to review new, easier way
  - Rewrote/fixed couple of functions
#### - 0.7.1 - 29.04.2019
  - Fixed Saving issue when file was in use and path was having dots in temporary path
#### - 0.7.0 - 17.04.2019
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
#### - 0.6.0 - 13.01.2019
  - Fixed merging of columns/cells in Table in circumstances with multiple paragraphs in a cell
  - Published with merging all files into single PSM1 and optimized PSD1 file which speeds up greatly Import-Module process
    If you want to read more about why I did it: <https://evotec.xyz/powershell-single-psm1-file-versus-multi-file-modules/>
#### - 0.5.17 - 13.01.2019
  - Fixed an empty chart problem in Word Blocks. If values/keys were empty it would create an empty chart preventing Word saving feature to work.
  - Fixed merging of columns/cells in Table
#### - 0.5.16 - 9.01.2019
  - Fixed Get-WordDocument support for path with special characters
#### - 0.5.15 - 29.12.2018
  - Added Add-WordPageCount with alias Add-WordPageNumber
  - Added some examples and tests for above
#### - 0.5.14 - 28.12.2018
  - Expanded support for Headers/Footers
  - Expanded Add-WordText with -Headers/Footers
  - Added some examples and tests
#### - 0.5.13 - 8.12.2018
  - Fix Set-WordTableRowMergeCells
  - Added Set-WordTableRowMergeCells TextMerge with Separator allowing for merging text from merged columns
  - Fix for Remove-WordText
  - Added tests for Set-WordTableRowMergeCells - does some additional testing to Tables just in case
#### - 0.5.12 - 8.12.2018
  - Fix for Set-WordTableRowMergeCells
#### - 0.5.11 - 30.11.2018
  - Added Alignment to Add-WordPicture
#### - 0.5.10 - 29.11.2018
  - Added Merge-WordDocument - which brings merging of Word Documents functionality
#### - 0.5.9 - 5.11.2018
  - Fix for Add-WordPicture
#### - 0.5.8 - 19.10.2018
  - Added small feature New-WordBlock for PSWinDocumentation
#### - 0.5.7 - 19.10.2018
  - Addition for Add-WordList to support singular (string,bool etc types)
  - Added Add-WordList Tests
#### - 0.5.6 - Bugfixes
#### - 0.5.4 - Added **NoLegend** switch to Charts (Add-WordPieChart, Add-WordLineChart, Add-WordBarChart)
#### - 0.5.1 - Small cleanup
#### - 0.5.0 - <https://evotec.xyz/pswriteword-version-0-5-1/>
#### - 0.4.7 - Added -ContinueFormatting to Add-WordText - same implementation as for Add-WordTable
#### - 0.4.6 - <https://evotec.xyz/pswriteword-add-wordtable-add-wordtext-expanded/>
#### - 0.4.1 - <https://evotec.xyz/pswriteword-updated-to-0-4-1-breaking-change-included/>