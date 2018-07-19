<#
/// <summary>
/// Custom property types.
/// </summary>
#>
Add-Type -TypeDefinition @"
public enum CustomPropertyType {
    Text,
    Date,
    NumberInteger,
    NumberDecimal,
    YesOrNo
}
"@
