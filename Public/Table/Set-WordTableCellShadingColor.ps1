function Set-WordTableCellShadingColor {
    [CmdletBinding()]
    param (
        [Xceed.Document.NET.InsertBeforeOrAfter] $Table,
        [nullable[int]] $RowNr,
        [nullable[int]] $ColumnNr,
        [nullable[System.Drawing.KnownColor]] $ShadingColor,
        [bool] $Supress = $false
    )
    if ($Table -ne $null -and $RowNr -ne $null -and $ColumnNr -ne $null -and $ShadingColor -ne $null) {
        $ConvertedColor = [System.Drawing.Color]::FromKnownColor($ShadingColor)
        $Cell = $Table.Rows[$RowNr].Cells[$ColumnNr]
        $Cell.Shading = $ConvertedColor
    }
    if ($Supress) { return } else { return $Table }
}



<#
 $Section3Table.Rows[0].Cells[0]
CELL

Paragraphs           : {Normal}
VerticalAlignment    : Center
Shading              : Color [White]
Width                : 115,5
MarginLeft           : NaN
MarginRight          : NaN
MarginTop            : NaN
MarginBottom         : NaN
FillColor            : Color [Empty]
TextDirection        : right
GridSpan             : 2
ParagraphsDeepSearch : {Normal}
Sections             : {}
Tables               : {}
Hyperlinks           : {}
Pictures             : {}
Lists                : {0}
Xml                  : <w:tc xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                         <w:tcPr>
                           <w:tcW w:w="2310" w:type="pct" />
                           <w:gridSpan w:val="2"></w:gridSpan>
                         </w:tcPr>
                         <w:p>
                           <w:pPr>
                             <w:jc w:val="center" />
                           </w:pPr>
                           <w:r>
                             <w:rPr>
                               <w:color w:val="0000FF"></w:color>
                             </w:rPr>
                             <w:t>Forest Summary</w:t>
                           </w:r>
                         </w:p>
                       </w:tc>
PackagePart          : System.IO.Packaging.ZipPackagePart
#>


<#

    /// <summary>
    /// Insert a row at the end of this table.
    /// </summary>
    /// <example>
    /// <code>
    /// // Load a document.
    /// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
    /// {
    ///     // Get the first table in this document.
    ///     Table table = document.Tables[0];
    ///
    ///     // Insert a new row at the end of this table.
    ///     Row row = table.InsertRow();
    ///
    ///     // Loop through each cell in this new row.
    ///     foreach (Cell c in row.Cells)
    ///     {
    ///         // Set the text of each new cell to "Hello".
    ///         c.Paragraphs[0].InsertText("Hello", false);
    ///     }
    ///
    ///     // Save the document to a new file.
    ///     document.SaveAs(@"C:\Example\Test2.docx");
    /// }// Release this document from memory.
    /// </code>
    /// </example>
    /// <returns>A new row.</returns>
        public Row InsertRow()
    {
      return this.InsertRow( this.RowCount );
    }

    /// <summary>
    /// Insert a copy of a row at the end of this table.
    /// </summary>
    /// <returns>A new row.</returns>
    public Row InsertRow( Row row, bool keepFormatting = false )
    {
      return this.InsertRow( row, this.RowCount, keepFormatting );
    }

    /// <summary>
    /// Insert a column to the right of a Table.
    /// </summary>
    /// <example>
    /// <code>
    /// // Load a document.
    /// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
    /// {
    ///     // Get the first Table in this document.
    ///     Table table = document.Tables[0];
    ///
    ///     // Insert a new column to this right of this table.
    ///     table.InsertColumn();
    ///
    ///     // Set the new columns text to "Row no."
    ///     table.Rows[0].Cells[table.ColumnCount - 1].Paragraph.InsertText("Row no.", false);
    ///
    ///     // Loop through each row in the table.
    ///     for (int i = 1; i &lt; table.Rows.Count; i++)
    ///     {
    ///         // The current row.
    ///         Row row = table.Rows[i];
    ///
    ///         // The cell in this row that belongs to the new column.
    ///         Cell cell = row.Cells[table.ColumnCount - 1];
    ///
    ///         // The first Paragraph that this cell houses.
    ///         Paragraph p = cell.Paragraphs[0];
    ///
    ///         // Insert this rows index.
    ///         p.InsertText(i.ToString(), false);
    ///     }
    ///
    ///     document.Save();
    /// }// Release this document from memory.
    /// </code>
    /// </example>




    /// <summary>
    /// Deletes a cell in a row and shift the others to the left.
    /// </summary>
    /// <param name="rowIndex">index of the row where a cell will be removed.</param>
    /// <param name="celIndex">index of the cell to remove in the row.</param>
    public void DeleteAndShiftCellsLeft( int rowIndex, int celIndex )
#>