<#
/// <summary>
/// Table Cell Border styles
/// source: http://msdn.microsoft.com/en-us/library/documentformat.openxml.wordprocessing.tablecellborders.aspx
/// </summary>
#>
Add-Type -TypeDefinition @"
public enum BorderStyle {
    Tcbs_none,
    Tcbs_single,
    Tcbs_thick,
    Tcbs_double,
    Tcbs_dotted,
    Tcbs_dashed,
    Tcbs_dotDash,
    Tcbs_dotDotDash,
    Tcbs_triple,
    Tcbs_thinThickSmallGap,
    Tcbs_thickThinSmallGap,
    Tcbs_thinThickThinSmallGap,
    Tcbs_thinThickMediumGap,
    Tcbs_thickThinMediumGap,
    Tcbs_thinThickThinMediumGap,
    Tcbs_thinThickLargeGap,
    Tcbs_thickThinLargeGap,
    Tcbs_thinThickThinLargeGap,
    Tcbs_wave,
    Tcbs_doubleWave,
    Tcbs_dashSmallGap,
    Tcbs_dashDotStroked,
    Tcbs_threeDEmboss,
    Tcbs_threeDEngrave,
    Tcbs_outset,
    Tcbs_inset,
    Tcbs_nil
}
"@