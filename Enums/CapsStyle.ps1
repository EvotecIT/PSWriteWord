<#
/// <summary>
/// Change the caps style of text, for use with Append and AppendLine.
/// </summary>
#>
enum CapsStyle {
    <#
    /// <summary>
    /// No caps, make all characters are lowercase.
    /// </summary>
    #>

    none
    <#
    /// <summary>
    /// All caps, make every character uppercase.
    /// </summary>
    #>
    caps
    <#
    /// <summary>
    /// Small caps, make all characters capital but with a small font size.
    /// </summary>
    #>
    smallCaps
}