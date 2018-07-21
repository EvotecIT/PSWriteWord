function Add-WordFooter {
    [CmdletBinding()]
    param (
        [Xceed.Words.NET.Container]$WordDocument,
        [nullable[bool]] $DifferentFirstPage,
        [nullable[bool]] $DifferentOddAndEvenPages,
        [bool] $Supress = $false
    )
    #$WordDocument.AddFooters()

    if ($DifferentOddAndEvenPages -ne $null ) { $WordDocument.DifferentFirstPage = $DifferentFirstPage }
    if ($DifferentOddAndEvenPages -ne $null ) { $WordDocument.DifferentOddAndEvenPages = $DifferentOddAndEvenPages }

    if ($Supress) { return } else { return $WordDocument.Footers }
}

function Add-WordHeader {
    [CmdletBinding()]
    param (
        [Xceed.Words.NET.Container]$WordDocument,
        [bool] $Supress = $false
    )
    $WordDocument.AddHeaders()
    if ($Supress) { return } else { return $WordDocument.Headers }
}

function Get-WordHeader {
    [CmdletBinding()]
    param (
        [Xceed.Words.NET.Container]$WordDocument,
        [bool] $Supress = $false
    )

}

function Get-WordFooter {
    [CmdletBinding()]
    param (
        [Xceed.Words.NET.Container]$WordDocument,
        [bool] $Supress = $false
    )

}

<#
    /// <summary>
    /// Returns a collection of Headers in this Document.
    /// A document typically contains three Headers.
    /// A default one (odd), one for the first page and one for even pages.
    /// </summary>
    /// <example>
    /// <code>
    /// // Create a document.
    /// using (DocX document = DocX.Create(@"Test.docx"))
    /// {
    ///    // Add header support to this document.
    ///    document.AddHeaders();
    ///
    ///    // Get a collection of all headers in this document.
    ///    Headers headers = document.Headers;
    ///
    ///    // The header used for the first page of this document.
    ///    Header first = headers.first;
    ///
    ///    // The header used for odd pages of this document.
    ///    Header odd = headers.odd;
    ///
    ///    // The header used for even pages of this document.
    ///    Header even = headers.even;
    /// }
    /// </code>
    /// </example>


        /// <summary>
    /// Returns a collection of Footers in this Document.
    /// A document typically contains three Footers.
    /// A default one (odd), one for the first page and one for even pages.
    /// </summary>
    /// <example>
    /// <code>
    /// // Create a document.
    /// using (DocX document = DocX.Create(@"Test.docx"))
    /// {
    ///    // Add footer support to this document.
    ///    document.AddFooters();
    ///
    ///    // Get a collection of all footers in this document.
    ///    Footers footers = document.Footers;
    ///
    ///    // The footer used for the first page of this document.
    ///    Footer first = footers.first;
    ///
    ///    // The footer used for odd pages of this document.
    ///    Footer odd = footers.odd;
    ///
    ///    // The footer used for even pages of this document.
    ///    Footer even = footers.even;
    /// }
    /// </code>
    /// </example>



    /// <summary>
    /// Should the Document use different Headers and Footers for odd and even pages?
    /// </summary>
    /// <example>
    /// <code>
    /// // Create a document.
    /// using (DocX document = DocX.Create(@"Test.docx"))
    /// {
    ///     // Add header support to this document.
    ///     document.AddHeaders();
    ///
    ///     // Get a collection of all headers in this document.
    ///     Headers headers = document.Headers;
    ///
    ///     // The header used for odd pages of this document.
    ///     Header odd = headers.odd;
    ///
    ///     // The header used for even pages of this document.
    ///     Header even = headers.even;
    ///
    ///     // Force the document to use a different header for odd and even pages.
    ///     document.DifferentOddAndEvenPages = true;
    ///
    ///     // Content can be added to the Headers in the same manor that it would be added to the main document.
    ///     Paragraph p1 = odd.InsertParagraph();
    ///     p1.Append("This is the odd pages header.");
    ///
    ///     Paragraph p2 = even.InsertParagraph();
    ///     p2.Append("This is the even pages header.");
    ///
    ///     // Save all changes to this document.
    ///     document.Save();
    /// }// Release this document from memory.
    /// </code>
    /// </example>


        /// <summary>
    /// Should the Document use an independent Header and Footer for the first page?
    /// </summary>
    /// <example>
    /// // Create a document.
    /// using (DocX document = DocX.Create(@"Test.docx"))
    /// {
    ///     // Add header support to this document.
    ///     document.AddHeaders();
    ///
    ///     // The header used for the first page of this document.
    ///     Header first = document.Headers.first;
    ///
    ///     // Force the document to use a different header for first page.
    ///     document.DifferentFirstPage = true;
    ///
    ///     // Content can be added to the Headers in the same manor that it would be added to the main document.
    ///     Paragraph p = first.InsertParagraph();
    ///     p.Append("This is the first pages header.");
    ///
    ///     // Save all changes to this document.
    ///     document.Save();
    /// }// Release this document from memory.
    /// </example>
#>
