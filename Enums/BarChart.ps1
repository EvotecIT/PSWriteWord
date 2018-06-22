<#
/// <summary>
/// Specifies the possible directions for a bar chart.
/// 21.2.3.3 ST_BarDir (Bar Direction)
/// </summary>
#>
enum BarDirection {
    Column
    Bar
}

<#
/// <summary>
/// Specifies the possible groupings for a bar chart.
/// 21.2.3.4 ST_BarGrouping (Bar Grouping)
/// </summary>
#>
enum BarGrouping {
    Clustered
    PercentStacked
    Stacked
    Standard
}
