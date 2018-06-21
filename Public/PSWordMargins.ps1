<#
    public static void Margins()
    {
      Console.WriteLine( "\tMargins()" );

      // Create a document.
      using( DocX document = DocX.Create( MarginSample.MarginSampleOutputDirectory + @"Margins.docx" ) )
      {
        // Add a title.
        document.InsertParagraph( "Document margins" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Set the page width to be smaller.
        document.PageWidth = 350f;

        // Set the document margins.
        document.MarginLeft = 85f;
        document.MarginRight = 85f;
        document.MarginTop = 0f;
        document.MarginBottom = 50f;

        // Add a paragraph. It will be affected by the document margins.
        var p = document.InsertParagraph("This is a paragraph from a document with a left margin of 85, a right margin of 85, a top margin of 0 and a bottom margin of 50.");

        document.Save();
        Console.WriteLine( "\tCreated: Margins.docx\n" );
      }
    }
#>

function Set-WordMargins {
    param(

    )
}