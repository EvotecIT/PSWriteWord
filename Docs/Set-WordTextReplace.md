---
external help file: PSWriteWord-help.xml
Module Name: PSWriteWord
online version:
schema: 2.0.0
---

# Set-WordTextReplace

## SYNOPSIS
Provides ability to search and replace certain words, phrases, or regular expressions in a word file.

## SYNTAX

```
Set-WordTextReplace [-Paragraph] <InsertBeforeOrAfter> [[-SearchValue] <String>] [[-ReplaceValue] <String>]
 [-TrackChanges] [[-RegexOptions] <RegexOptions>] [[-NewFormatting] <Formatting>]
 [[-MatchFormatting] <Formatting>] [[-MatchFormattingOptions] <MatchFormattingOptions>] [-EscapeRegEx]
 [-UseRegExSubstitutions] [-RemoveEmptyParagraph] [[-Suppress] <Boolean>] [<CommonParameters>]
```

## DESCRIPTION
Provides ability to search and replace certain words, phrases, or regular expressions in a word file.

## EXAMPLES

### EXAMPLE 1
```
$FilePath = "C:\Users\przemyslaw.klys\OneDrive - Evotec\Desktop\Word.docx"
```

$FilePath1 = "C:\Users\przemyslaw.klys\OneDrive - Evotec\Desktop\Word1.docx"
$doc = Get-WordDocument -FilePath $FilePath
$word = "Sample"
$formatObj = New-Object Xceed.Document.NET.Formatting
$formatObj.FontColor = "Red"
foreach ($p in $doc.Paragraphs) {
    Set-WordTextReplace -Paragraph $p -SearchValue $word -ReplaceValue $word -NewFormatting $formatObj -Supress $false
}
Save-WordDocument -Document $doc -FilePath $FilePath1

## PARAMETERS

### -Paragraph
Provide paragraph to search for text

```yaml
Type: InsertBeforeOrAfter
Parameter Sets: (All)
Aliases:

Required: True
Position: 1
Default value: None
Accept pipeline input: True (ByPropertyName, ByValue)
Accept wildcard characters: False
```

### -SearchValue
Value to search for

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 2
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ReplaceValue
Value to replace with

```yaml
Type: String
Parameter Sets: (All)
Aliases: NewValue

Required: False
Position: 3
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TrackChanges
Track changes, default is off

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -RegexOptions
The regex options to use when searching for the search value.
Default is none.

```yaml
Type: RegexOptions
Parameter Sets: (All)
Aliases:
Accepted values: None, IgnoreCase, Multiline, ExplicitCapture, Compiled, Singleline, IgnorePatternWhitespace, RightToLeft, ECMAScript, CultureInvariant

Required: False
Position: 4
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NewFormatting
The formatting to apply to the text being inserted.

```yaml
Type: Formatting
Parameter Sets: (All)
Aliases:

Required: False
Position: 5
Default value: [Xceed.Document.NET.Formatting]::new()
Accept pipeline input: False
Accept wildcard characters: False
```

### -MatchFormatting
The formatting that the text must match in order to be replaced.

```yaml
Type: Formatting
Parameter Sets: (All)
Aliases:

Required: False
Position: 6
Default value: [Xceed.Document.NET.Formatting]::new()
Accept pipeline input: False
Accept wildcard characters: False
```

### -MatchFormattingOptions
How should formatting be matched?
ExactMatch (default) or SubsetMatch

```yaml
Type: MatchFormattingOptions
Parameter Sets: (All)
Aliases:
Accepted values: ExactMatch, SubsetMatch

Required: False
Position: 7
Default value: ExactMatch
Accept pipeline input: False
Accept wildcard characters: False
```

### -EscapeRegEx
True if the oldValue needs to be escaped, otherwise false.
If it represents a valid RegEx pattern this should be false.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -UseRegExSubstitutions
True if RegEx-like replace should be performed, i.e.
if newValue contains RegEx substitutions.
Does not perform named-group substitutions (only numbered groups).

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -RemoveEmptyParagraph
Remove empty paragraph

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -Suppress
{{ Fill Suppress Description }}

```yaml
Type: Boolean
Parameter Sets: (All)
Aliases: Supress

Required: False
Position: 8
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

## OUTPUTS

## NOTES
General notes

## RELATED LINKS
