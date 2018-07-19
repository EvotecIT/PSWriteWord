<#
/// <summary>
/// Paragraph edit types
/// </summary>
#>
Add-Type -TypeDefinition @"
public enum EditType {
    ins,
    del
}
"@