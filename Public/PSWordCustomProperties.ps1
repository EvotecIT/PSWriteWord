<#
      using( DocX document = DocX.Create( DocumentSample.DocumentSampleOutputDirectory + @"AddCustomProperties.docx" ) )
      {
        // Add a title
        document.InsertParagraph( "Adding Custom Properties to a document" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        //Add custom properties to document.
        document.AddCustomProperty( new CustomProperty( "CompanyName", "Xceed Software inc." ) );
        document.AddCustomProperty( new CustomProperty( "Product", "Xceed Words for .NET" ) );
        document.AddCustomProperty( new CustomProperty( "Address", "10 Boul. de Mortagne" ) );
        document.AddCustomProperty( new CustomProperty( "Date", DateTime.Now ) );

        // Add a paragraph displaying the number of custom properties.
        var p = document.InsertParagraph( "This document contains " ).Append( document.CustomProperties.Count.ToString() ).Append(" Custom Properties :");
        p.SpacingAfter( 30 );

        // Display each propertie's name and value.
        foreach( var prop in document.CustomProperties )
        {
          document.InsertParagraph( prop.Value.Name ).Append( " = " ).Append( prop.Value.Value.ToString() ).AppendLine();
        }

        // Save this document to disk.
        document.Save();
        Console.WriteLine( "\tCreated: AddCustomProperties.docx\n" );
      }
#>
function Add-WordCustomProperty {
    [CmdletBinding()]
    param (
        [Xceed.Words.NET.Container]$WordDocument,
        [string] $Name,
        [string] $Value
    )
    $CustomProperty = New-Object -TypeName Xceed.Words.NET.CustomProperty -ArgumentList $Name, $Value
    $WordDocument.AddCustomProperty($CustomProperty)
}

function Get-WordCustomProperty {
    [CmdletBinding()]
    param (
        [Xceed.Words.NET.Container]$WordDocument,
        [string] $Name
    )
    if ($Property -eq $null) {
        return $WordDocument.CustomProperties.Values
    } else {
        return $WordDocument.CustomProperties.$Name.Value
    }
}