<#
/// <summary>
/// Specifies the kind of grouping for a column, line, or area chart.
/// 21.2.2.76 grouping (Grouping)
/// </summary>
#>
Add-Type -TypeDefinition @"
public enum Grouping {
    PercentStacked,
    Stacked,
    Standard
}
"@