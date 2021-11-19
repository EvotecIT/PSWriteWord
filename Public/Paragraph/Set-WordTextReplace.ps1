function Set-WordTextReplace {
    <#
    .SYNOPSIS
    Provides ability to search and replace certain words, phrases, or regular expressions in a word file.

    .DESCRIPTION
    Provides ability to search and replace certain words, phrases, or regular expressions in a word file.

    .PARAMETER Paragraph
    Provide paragraph to search for text

    .PARAMETER SearchValue
    Value to search for

    .PARAMETER ReplaceValue
    Value to replace with

    .PARAMETER TrackChanges
    Track changes, default is off

    .PARAMETER RegexOptions
    The regex options to use when searching for the search value. Default is none.

    .PARAMETER NewFormatting
    The formatting to apply to the text being inserted.

    .PARAMETER MatchFormatting
    The formatting that the text must match in order to be replaced.

    .PARAMETER MatchFormattingOptions
    How should formatting be matched? ExactMatch (default) or SubsetMatch

    .PARAMETER escapeRegEx
    True if the oldValue needs to be escaped, otherwise false. If it represents a valid RegEx pattern this should be false.

    .PARAMETER useRegExSubstitutions
    True if RegEx-like replace should be performed, i.e. if newValue contains RegEx substitutions. Does not perform named-group substitutions (only numbered groups).

    .PARAMETER removeEmptyParagraph
    Remove empty paragraph

    .EXAMPLE
    $FilePath = "C:\Users\przemyslaw.klys\OneDrive - Evotec\Desktop\Word.docx"
    $FilePath1 = "C:\Users\przemyslaw.klys\OneDrive - Evotec\Desktop\Word1.docx"
    $doc = Get-WordDocument -FilePath $FilePath
    $word = "Sample"
    $formatObj = New-Object Xceed.Document.NET.Formatting
    $formatObj.FontColor = "Red"
    foreach ($p in $doc.Paragraphs) {
        Set-WordTextReplace -Paragraph $p -SearchValue $word -ReplaceValue $word -NewFormatting $formatObj -Supress $false
    }
    Save-WordDocument -Document $doc -FilePath $FilePath1

    .NOTES
    General notes
    #>
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline, Mandatory)][Xceed.Document.NET.InsertBeforeOrAfter] $Paragraph,
        [string] $SearchValue,
        [alias('NewValue')][string] $ReplaceValue,
        [switch] $TrackChanges,
        [System.Text.RegularExpressions.RegexOptions] $RegexOptions = [System.Text.RegularExpressions.RegexOptions]::None,
        [Xceed.Document.NET.Formatting] $NewFormatting = [Xceed.Document.NET.Formatting]::new(),
        [Xceed.Document.NET.Formatting] $MatchFormatting = [Xceed.Document.NET.Formatting]::new(),
        [Xceed.Document.NET.MatchFormattingOptions] $MatchFormattingOptions = [Xceed.Document.NET.MatchFormattingOptions]::ExactMatch,
        [switch] $EscapeRegEx,
        [switch] $UseRegExSubstitutions,
        [switch] $RemoveEmptyParagraph,
        [alias('Supress')][bool] $Suppress = $false
    )
    #void ReplaceText(string searchValue, string newValue, bool trackChanges, System.Text.RegularExpressions.RegexOptions options, Xceed.Document.NET.Formatting newFormatting, Xceed.Document.NET.Formatting matchFormatting, Xceed.Document.NET.MatchFormattingOptions fo, bool escapeRegEx, bool useRegExSubstitutions, bool removeEmptyParagraph)
    #void ReplaceText(string findPattern, System.Func[string,string] regexMatchHandler, bool trackChanges, System.Text.RegularExpressions.RegexOptions options, Xceed.Document.NET.Formatting newFormatting, Xceed.Document.NET.Formatting matchFormatting, Xceed.Document.NET.MatchFormattingOptions fo, bool removeEmptyParagraph)
    if ($Paragraph) {
        $Paragraph = $Paragraph.ReplaceText($SearchValue, $ReplaceValue, $TrackChanges.IsPresent, $RegexOptions, $NewFormatting, $matchFormatting, $MatchFormattingOptions, $EscapeRegEx.IsPresent, $UseRegExSubstitutions.IsPresent, $RemoveEmptyParagraph.IsPresent)
        if ($Suppress) { return } else { return $Paragraph }
    }
}