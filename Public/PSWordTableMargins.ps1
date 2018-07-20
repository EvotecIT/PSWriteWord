<#



    // <summary>
    // Gets or Sets this Cells vertical alignment.
    // </summary>
    // <example>
    // Creates a table with 3 cells and sets the vertical alignment of each to 1 of the 3 available options.
    // <code>
    // Create a new document.
    //using(DocX document = DocX.Create("Test.docx"))
    //{
    //    // Insert a Table into this document.
    //    Table t = document.InsertTable(3, 1);
    //
    //    // Set the design of the Table such that we can easily identify cell boundaries.
    //    t.Design = TableDesign.TableGrid;
    //
    //    // Set the height of the row bigger than default.
    //    // We need to be able to see the difference in vertical cell alignment options.
    //    t.Rows[0].Height = 100;
    //
    //    // Set the vertical alignment of cell0 to top.
    //    Cell c0 = t.Rows[0].Cells[0];
    //    c0.InsertParagraph("VerticalAlignment.Top");
    //    c0.VerticalAlignment = VerticalAlignment.Top;
    //
    //    // Set the vertical alignment of cell1 to center.
    //    Cell c1 = t.Rows[0].Cells[1];
    //    c1.InsertParagraph("VerticalAlignment.Center");
    //    c1.VerticalAlignment = VerticalAlignment.Center;
    //
    //    // Set the vertical alignment of cell2 to bottom.
    //    Cell c2 = t.Rows[0].Cells[2];
    //    c2.InsertParagraph("VerticalAlignment.Bottom");
    //    c2.VerticalAlignment = VerticalAlignment.Bottom;
    //
    //    // Save the document.
    //    document.Save();
    //}
    // </code>
    // </example>



        /// <summary>
    /// LeftMargin in pixels.
    /// </summary>
    /// <example>
    /// <code>
    /// // Create a new document.
    ///using (DocX document = DocX.Create("Test.docx"))
    ///{
    ///    // Insert table into this document.
    ///    Table t = document.InsertTable(3, 3);
    ///    t.Design = TableDesign.TableGrid;
    ///
    ///    // Get the center cell.
    ///    Cell center = t.Rows[1].Cells[1];
    ///
    ///    // Insert some text so that we can see the effect of the Margins.
    ///    center.Paragraphs[0].Append("Center Cell");
    ///
    ///    // Set the center cells Left, Margin to 10.
    ///    center.MarginLeft = 25;
    ///
    ///    // Save the document.
    ///    document.Save();
    ///}
    /// </code>
    /// </example>



    /// <summary>
    /// RightMargin in pixels.
    /// </summary>
    /// <example>
    /// <code>
    /// // Create a new document.
    ///using (DocX document = DocX.Create("Test.docx"))
    ///{
    ///    // Insert table into this document.
    ///    Table t = document.InsertTable(3, 3);
    ///    t.Design = TableDesign.TableGrid;
    ///
    ///    // Get the center cell.
    ///    Cell center = t.Rows[1].Cells[1];
    ///
    ///    // Insert some text so that we can see the effect of the Margins.
    ///    center.Paragraphs[0].Append("Center Cell");
    ///
    ///    // Set the center cells Right, Margin to 10.
    ///    center.MarginRight = 25;
    ///
    ///    // Save the document.
    ///    document.Save();
    ///}
    /// </code>
    /// </example>



    /// <summary>
    /// Merge cells in given column starting with startRow and ending with endRow.
    /// </summary>
    public void MergeCellsInColumn( int columnIndex, int startRow, int endRow )


        /// <summary>
    /// Remove this Table from this document.
    /// </summary>
    /// <example>
    /// Remove the first Table from this document.
    /// <code>
    /// // Load a document into memory.
    /// using (DocX document = DocX.Load(@"Test.docx"))
    /// {
    ///     // Get the first Table in this document.
    ///     Table t = d.Tables[0];
    ///
    ///     // Remove this Table.
    ///     t.Remove();
    ///
    ///     // Save all changes made to the document.
    ///     document.Save();
    /// } // Release this document from memory.
    /// </code>
    /// </example>


        /// <summary>
    /// Merge cells starting with startIndex and ending with endIndex.
    /// </summary>
    public void MergeCells( int startIndex, int endIndex )
#>