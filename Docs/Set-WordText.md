---
external help file: PSWriteWord-help.xml
Module Name: PSWriteWord
online version:
schema: 2.0.0
---

# Set-WordText

## SYNOPSIS
{{Fill in the Synopsis}}

## SYNTAX

```
Set-WordText [[-Paragraph] <InsertBeforeOrAfter[]>] [[-Color] <Color[]>] [[-FontSize] <Double[]>]
 [[-FontFamily] <String[]>] [[-Bold] <Nullable`1[]>] [[-Italic] <Nullable`1[]>]
 [[-UnderlineStyle] <UnderlineStyle[]>] [[-UnderlineColor] <Color[]>] [[-SpacingAfter] <Double[]>]
 [[-SpacingBefore] <Double[]>] [[-Spacing] <Double[]>] [[-Highlight] <Highlight[]>]
 [[-CapsStyle] <CapsStyle[]>] [[-StrikeThrough] <StrikeThrough[]>] [[-HeadingType] <HeadingType[]>]
 [[-PercentageScale] <Int32[]>] [[-Misc] <Misc[]>] [[-Language] <String[]>] [[-Kerning] <Int32[]>]
 [[-Hidden] <Nullable`1[]>] [[-Position] <Int32[]>] [[-NewLine] <Nullable`1[]>] [-KeepLinesTogether]
 [-KeepWithNextParagraph] [[-IndentationFirstLine] <Single[]>] [[-IndentationHanging] <Single[]>]
 [[-Alignment] <Alignment[]>] [[-Direction] <Direction[]>] [[-ShadingType] <ShadingType[]>]
 [[-Script] <Script[]>] [[-Supress] <Boolean>] [<CommonParameters>]
```

## DESCRIPTION
{{Fill in the Description}}

## EXAMPLES

### Example 1
```powershell
PS C:\> {{ Add example code here }}
```

{{ Add example description here }}

## PARAMETERS

### -Alignment
{{Fill Alignment Description}}

```yaml
Type: Alignment[]
Parameter Sets: (All)
Aliases:
Accepted values: left, center, right, both

Required: False
Position: 24
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Bold
{{Fill Bold Description}}

```yaml
Type: Nullable`1[]
Parameter Sets: (All)
Aliases: B

Required: False
Position: 4
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -CapsStyle
{{Fill CapsStyle Description}}

```yaml
Type: CapsStyle[]
Parameter Sets: (All)
Aliases: CA
Accepted values: none, caps, smallCaps

Required: False
Position: 12
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Color
{{Fill Color Description}}

```yaml
Type: Color[]
Parameter Sets: (All)
Aliases: C

Required: False
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Direction
{{Fill Direction Description}}

```yaml
Type: Direction[]
Parameter Sets: (All)
Aliases:
Accepted values: LeftToRight, RightToLeft

Required: False
Position: 25
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FontFamily
{{Fill FontFamily Description}}

```yaml
Type: String[]
Parameter Sets: (All)
Aliases: FontName

Required: False
Position: 3
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FontSize
{{Fill FontSize Description}}

```yaml
Type: Double[]
Parameter Sets: (All)
Aliases: S

Required: False
Position: 2
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -HeadingType
{{Fill HeadingType Description}}

```yaml
Type: HeadingType[]
Parameter Sets: (All)
Aliases: HT
Accepted values: Heading1, Heading2, Heading3, Heading4, Heading5, Heading6, Heading7, Heading8, Heading9

Required: False
Position: 14
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Hidden
{{Fill Hidden Description}}

```yaml
Type: Nullable`1[]
Parameter Sets: (All)
Aliases:

Required: False
Position: 19
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Highlight
{{Fill Highlight Description}}

```yaml
Type: Highlight[]
Parameter Sets: (All)
Aliases: H
Accepted values: yellow, green, cyan, magenta, blue, red, darkBlue, darkCyan, darkGreen, darkMagenta, darkRed, darkYellow, darkGray, lightGray, black, none

Required: False
Position: 11
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IndentationFirstLine
{{Fill IndentationFirstLine Description}}

```yaml
Type: Single[]
Parameter Sets: (All)
Aliases:

Required: False
Position: 22
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IndentationHanging
{{Fill IndentationHanging Description}}

```yaml
Type: Single[]
Parameter Sets: (All)
Aliases:

Required: False
Position: 23
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Italic
{{Fill Italic Description}}

```yaml
Type: Nullable`1[]
Parameter Sets: (All)
Aliases: I

Required: False
Position: 5
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -KeepLinesTogether
{{Fill KeepLinesTogether Description}}

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -KeepWithNextParagraph
{{Fill KeepWithNextParagraph Description}}

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Kerning
{{Fill Kerning Description}}

```yaml
Type: Int32[]
Parameter Sets: (All)
Aliases:

Required: False
Position: 18
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Language
{{Fill Language Description}}

```yaml
Type: String[]
Parameter Sets: (All)
Aliases:

Required: False
Position: 17
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Misc
{{Fill Misc Description}}

```yaml
Type: Misc[]
Parameter Sets: (All)
Aliases:
Accepted values: none, shadow, outline, outlineShadow, emboss, engrave

Required: False
Position: 16
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NewLine
{{Fill NewLine Description}}

```yaml
Type: Nullable`1[]
Parameter Sets: (All)
Aliases:

Required: False
Position: 21
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Paragraph
{{Fill Paragraph Description}}

```yaml
Type: InsertBeforeOrAfter[]
Parameter Sets: (All)
Aliases:

Required: False
Position: 0
Default value: None
Accept pipeline input: True (ByPropertyName, ByValue)
Accept wildcard characters: False
```

### -PercentageScale
{{Fill PercentageScale Description}}

```yaml
Type: Int32[]
Parameter Sets: (All)
Aliases:

Required: False
Position: 15
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Position
{{Fill Position Description}}

```yaml
Type: Int32[]
Parameter Sets: (All)
Aliases:

Required: False
Position: 20
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Script
{{Fill Script Description}}

```yaml
Type: Script[]
Parameter Sets: (All)
Aliases:
Accepted values: superscript, subscript, none

Required: False
Position: 27
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ShadingType
{{Fill ShadingType Description}}

```yaml
Type: ShadingType[]
Parameter Sets: (All)
Aliases:
Accepted values: Text, Paragraph

Required: False
Position: 26
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Spacing
{{Fill Spacing Description}}

```yaml
Type: Double[]
Parameter Sets: (All)
Aliases: SP

Required: False
Position: 10
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SpacingAfter
{{Fill SpacingAfter Description}}

```yaml
Type: Double[]
Parameter Sets: (All)
Aliases: SA

Required: False
Position: 8
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SpacingBefore
{{Fill SpacingBefore Description}}

```yaml
Type: Double[]
Parameter Sets: (All)
Aliases: SB

Required: False
Position: 9
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -StrikeThrough
{{Fill StrikeThrough Description}}

```yaml
Type: StrikeThrough[]
Parameter Sets: (All)
Aliases: ST
Accepted values: none, strike, doubleStrike

Required: False
Position: 13
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Supress
{{Fill Supress Description}}

```yaml
Type: Boolean
Parameter Sets: (All)
Aliases:

Required: False
Position: 28
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -UnderlineColor
{{Fill UnderlineColor Description}}

```yaml
Type: Color[]
Parameter Sets: (All)
Aliases: UC

Required: False
Position: 7
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -UnderlineStyle
{{Fill UnderlineStyle Description}}

```yaml
Type: UnderlineStyle[]
Parameter Sets: (All)
Aliases: U
Accepted values: none, singleLine, words, doubleLine, dotted, thick, dash, dotDash, dotDotDash, wave, dottedHeavy, dashedHeavy, dashDotHeavy, dashDotDotHeavy, dashLongHeavy, dashLong, wavyDouble, wavyHeavy

Required: False
Position: 6
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see about_CommonParameters (http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

### Xceed.Words.NET.InsertBeforeOrAfter

## OUTPUTS

### System.Object

## NOTES

## RELATED LINKS
