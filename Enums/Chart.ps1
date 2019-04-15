Add-Type -TypeDefinition @"
public enum ChartLegendPosition {
    Top,
    Bottom,
    Left,
    Right,
    TopRight
}
"@
Add-Type -TypeDefinition @"
public enum DisplayBlanksAs {
    Gap,
    Span,
    Zero
}
"@