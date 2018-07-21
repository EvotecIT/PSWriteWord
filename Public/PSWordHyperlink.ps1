<#
      using( DocX document = DocX.Create( HyperlinkSample.HyperlinkSampleOutputDirectory + @"Hyperlinks.docx" ) )
      {
        // Add a title
        document.InsertParagraph( "Insert/Remove Hyperlinks" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Add an Hyperlink into this document.
        var h = document.AddHyperlink( "google", new Uri( "http://www.google.com" ) );

        // Add a paragraph.
        var p = document.InsertParagraph( "The  hyperlink has been inserted in this paragraph." );
        // insert an hyperlink at specific index in this paragraph.
        p.InsertHyperlink( h, 4 );
        p.SpacingAfter( 40d );

        // Get the first hyperlink in the document.
        var hyperlink = document.Hyperlinks.FirstOrDefault();
        if( hyperlink != null )
        {
          // Modify its text and Uri.
          hyperlink.Text = "xceed";
          hyperlink.Uri = new Uri( "http://www.xceed.com/" );
        }

        // Add an Hyperlink to this document.
        var h2 = document.AddHyperlink( "xceed", new Uri( "http://www.xceed.com/" ) );
        // Add a paragraph.
        var p2 = document.InsertParagraph( "A formatted hyperlink has been added at the end of this paragraph: " );
        // Append an hyperlink to a paragraph.
        p2.AppendHyperlink( h2 ).Color( Color.Blue ).UnderlineStyle( UnderlineStyle.singleLine );
        p2.Append( "." ).SpacingAfter( 40d );

        // Add an Hyperlink to this document.
        var h3 = document.AddHyperlink( "microsoft", new Uri( "http://www.microsoft.com" ) );
        // Add a paragraph
        var p3 = document.InsertParagraph( "The hyperlink from this paragraph has been removed. " );
        // Append an hyperlink to a paragraph.
        p3.AppendHyperlink( h3 ).Color( Color.Green ).UnderlineStyle( UnderlineStyle.singleLine ).Italic();

        // Remove the first hyperlink of paragraph 3.
        p3.RemoveHyperlink( 0 );

        document.Save();
        Console.WriteLine( "\tCreated: Hyperlinks.docx\n" );
      }
#>

function Add-WordHyperLink {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.Container]$WordDocument,
        [string] $UrlText,
        [string] $UrlLink,
        [bool] $Supress = $false
    )
    $Url = New-Object -TypeName Uri -ArgumentList $UrlLink

    return $WordDocument.AddHyperlink( $UrlText, $Url )
}
function Set-WordHyperLink {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)] [Xceed.Words.NET.Container]$WordDocument,
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)] [Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [Xceed.Words.NET.DocXElement] $Value,
        [bool] $Supress = $false
    )
    $Data = $Paragraph.InsertHyperlink($Value)

    if ($Supress -eq $false) { return $Data } else { return }

}