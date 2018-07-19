Add-Type -TypeDefinition @"
public enum ChartLegendPosition {
    Top,
    Bottom,
    Left,
    Right,
    TopRight
}
"@

<#
/// <summary>
/// Specifies the possible ways to display blanks.
/// 21.2.3.10 ST_DispBlanksAs (Display Blanks As)
/// </summary>
#>
Add-Type -TypeDefinition @"
public enum DisplayBlanksAs {
    Gap,
    Span,
    Zero
}
"@