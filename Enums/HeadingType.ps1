<#
    // The following headings appear in the same list in Word but they do not work in the same way (they are character based headings not paragraph based headings)
    // NoSpacing
    // Title Subtitle
    // Quote IntenseQuote
    // Emphasis IntenseEmphasis
    // Strong
    // ListParagraph
    // SubtleReference IntenseReference
    // BookTitle
#>
Add-Type -TypeDefinition @"
public enum HeadingType {
    Heading1,
    Heading2,
    Heading3,
    Heading4,
    Heading5,
    Heading6,
    Heading7,
    Heading8,
    Heading9
}
"@