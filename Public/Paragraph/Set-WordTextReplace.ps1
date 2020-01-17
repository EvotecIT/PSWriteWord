function Set-WordTextReplace {
    <#
    .SYNOPSIS
    Short description

    .DESCRIPTION
    Long description

    .PARAMETER Paragraph
    Parameter description

    .PARAMETER SearchValue
    Parameter description

    .PARAMETER ReplaceValue
    Parameter description

    .PARAMETER TrackChanges
    Track changes

    .PARAMETER RegexOptions
    Parameter description

    .PARAMETER NewFormatting
    The formatting to apply to the text being inserted.

    .PARAMETER MatchFormatting
    The formatting that the text must match in order to be replaced.

    .PARAMETER MatchFormattingOptions
    How should formatting be matched?

    .PARAMETER escapeRegEx
    True if the oldValue needs to be escaped, otherwise false. If it represents a valid RegEx pattern this should be false.

    .PARAMETER useRegExSubstitutions
    True if RegEx-like replace should be performed, i.e. if newValue contains RegEx substitutions. Does not perform named-group substitutions (only numbered groups).

    .PARAMETER removeEmptyParagraph
    Remove empty paragraph

    .EXAMPLE
    An example

    .NOTES
    General notes
    #>
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Document.NET.InsertBeforeOrAfter] $Paragraph,
        [string] $SearchValue,
        [string] $ReplaceValue,
        [bool] $TrackChanges,
        [System.Text.RegularExpressions.RegexOptions] $RegexOptions,
        [Xceed.Document.NET.Formatting] $NewFormatting,
        [Xceed.Document.NET.Formatting] $MatchFormatting,
        [Xceed.Document.NET.MatchFormattingOptions] $MatchFormattingOptions,
        [bool] $escapeRegEx = $true,
        [bool] $useRegExSubstitutions = $false,
        [bool] $removeEmptyParagraph = $true
    )
    #void ReplaceText(string searchValue, string newValue, bool trackChanges, System.Text.RegularExpressions.RegexOptions options, Xceed.Document.NET.Formatting newFormatting, Xceed.Document.NET.Formatting matchFormatting, Xceed.Document.NET.MatchFormattingOptions fo, bool escapeRegEx, bool useRegExSubstitutions, bool removeEmptyParagraph)
    #void ReplaceText(string findPattern, System.Func[string,string] regexMatchHandler, bool trackChanges, System.Text.RegularExpressions.RegexOptions options, Xceed.Document.NET.Formatting newFormatting, Xceed.Document.NET.Formatting matchFormatting, Xceed.Document.NET.MatchFormattingOptions fo, bool removeEmptyParagraph)

    $Paragraph.ReplaceText
    #$Paragraph.ReplaceText($SearchValue, $ReplaceValue, $TrackChanges, $RegexOptions, $NewFormatting, $MatchFormatting, $MatchFormattingOptions, $escapeRegEx, $useRegExSubstitutions, $removeEmptyParagraph)
}

#Set-WordTextReplace