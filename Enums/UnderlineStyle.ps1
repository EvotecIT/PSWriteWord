Add-Type -TypeDefinition @"
public enum UnderlineStyle {
    none = 0,
    singleLine = 1,
    words = 2,
    doubleLine = 3,
    dotted = 4,
    thick = 6,
    dash = 7,
    dotDash = 9,
    dotDotDash = 10,
    wave = 11,
    dottedHeavy = 20,
    dashedHeavy = 23,
    dashDotHeavy = 25,
    dashDotDotHeavy = 26,
    dashLongHeavy = 27,
    dashLong = 39,
    wavyDouble = 43,
    wavyHeavy = 55
}
"@