

<#
    public static List<ChartData> CreateCanadaExpenses()
    {
      var canada = new List<ChartData>();
      canada.Add( new ChartData() { Category = "Food", Expenses = 100 } );
      canada.Add( new ChartData() { Category = "Housing", Expenses = 120 } );
      canada.Add( new ChartData() { Category = "Transportation", Expenses = 140 } );
      canada.Add( new ChartData() { Category = "Health Care", Expenses = 150 } );
      return canada;
    }

    public static List<ChartData> CreateUSAExpenses()
    {
      var usa = new List<ChartData>();
      usa.Add( new ChartData() { Category = "Food", Expenses = 200 } );
      usa.Add( new ChartData() { Category = "Housing", Expenses = 150 } );
      usa.Add( new ChartData() { Category = "Transportation", Expenses = 110 } );
      usa.Add( new ChartData() { Category = "Health Care", Expenses = 100 } );
      return usa;
    }

    public static List<ChartData> CreateBrazilExpenses()
    {
      var brazil = new List<ChartData>();
      brazil.Add( new ChartData() { Category = "Food", Expenses = 125 } );
      brazil.Add( new ChartData() { Category = "Housing", Expenses = 80 } );
      brazil.Add( new ChartData() { Category = "Transportation", Expenses = 110 } );
      brazil.Add( new ChartData() { Category = "Health Care", Expenses = 60 } );
      return brazil;
    }

     using( DocX document = DocX.Create( ChartSample.ChartSampleOutputDirectory + @"PieChart.docx" ) )
      {
        // Add a title
        document.InsertParagraph( "Pie Chart" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Create a pie chart.
        var c = new PieChart();
        c.AddLegend( ChartLegendPosition.Left, false );

        // Create the data.
        var brazil = ChartData.CreateBrazilExpenses();

        // Create and add series
        var s1 = new Series( "Canada" );
        s1.Bind( brazil, "Category", "Expenses" );
        c.AddSeries( s1 );

        // Insert chart into document
        document.InsertParagraph( "Expenses(M$) for selected categories of Canada" ).FontSize( 15 ).SpacingAfter( 10d );
        document.InsertChart( c );

        document.Save();
        Console.WriteLine( "\tCreated: PieChart.docx\n" );
      }
    }
#>

Import-Module PSWriteWord #-Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-CreateCharts1.docx"

$WordDocument = New-WordDocument $FilePath
Add-WordText -WordDocument $WordDocument -Text 'This is text that has font size of 15', ' and this is font size of 10 ', ' while this will be 12.' `
    -FontSize 15, 10 `
    -Color Blue, Red `
    -Bold $true, $false, $true `
    -Italic $true, $true

$Test0 = @{
    Expenses = 100
    Category = 'Food'
}
$Test1 = @{
    Expenses = 200
    Category = 'Food'
}

Class Car {
    [String]$vin
    [int]$numberOfWheels = 4
}

$Test = New-Object car
$Test.Vin = 'test'
$Test.numberOfWheels = 4


$List1 = New-ArrayList
Add-ToArray -List $List1 -Element $Test
#Add-ToArray -List $List1 -Element $Test

[Xceed.Words.NET.Series] $series = New-Object -TypeName Xceed.Words.NET.Series -ArgumentList 'Name'
$series.Color = [System.Drawing.Color]::Aqua
$series.Bind($List1, "Vin", "NUMBEROFWHEELS")

[Xceed.Words.NET.PieChart] $chart = New-Object -TypeName Xceed.Words.NET.PieChart
$chart.AddLegend('left', $true)
$chart.AddSeries($series)
$chart.

$WordDocument.InsertChart($chart)


Save-WordDocument $WordDocument