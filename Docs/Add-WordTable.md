---
external help file: PSWriteWord-help.xml
Module Name: PSWriteWord
online version:
schema: 2.0.0
---

# Add-WordTable

## SYNOPSIS
{{Fill in the Synopsis}}

## SYNTAX

```
Add-WordTable [[-WordDocument] <Container>] [[-Paragraph] <InsertBeforeOrAfter>]
 [[-Table] <InsertBeforeOrAfter>] [[-DataTable] <Object>] [[-AutoFit] <AutoFit>] [[-Design] <TableDesign>]
 [[-Direction] <Direction>] [-BreakPageAfterTable] [-BreakPageBeforeTable] [[-BreakAcrossPages] <Boolean>]
 [[-MaximumColumns] <Int32>] [[-Titles] <String[]>] [-DoNotAddTitle] [[-ColummnWidth] <Single[]>]
 [[-TableWidth] <Single>] [[-Percentage] <Boolean>] [[-Color] <Color[]>] [[-FontSize] <Double[]>]
 [[-FontFamily] <String[]>] [[-Bold] <Nullable`1[]>] [[-Italic] <Nullable`1[]>]
 [[-UnderlineStyle] <UnderlineStyle[]>] [[-UnderlineColor] <Color[]>] [[-SpacingAfter] <Double[]>]
 [[-SpacingBefore] <Double[]>] [[-Spacing] <Double[]>] [[-Highlight] <Highlight[]>]
 [[-CapsStyle] <CapsStyle[]>] [[-StrikeThrough] <StrikeThrough[]>] [[-HeadingType] <HeadingType[]>]
 [[-PercentageScale] <Int32[]>] [[-Misc] <Misc[]>] [[-Language] <String[]>] [[-Kerning] <Int32[]>]
 [[-Hidden] <Nullable`1[]>] [[-Position] <Int32[]>] [[-IndentationFirstLine] <Single[]>]
 [[-IndentationHanging] <Single[]>] [[-Alignment] <Alignment[]>] [[-DirectionFormatting] <Direction[]>]
 [[-ShadingType] <ShadingType[]>] [[-Script] <Script[]>] [[-NewLine] <Nullable`1[]>] [-KeepLinesTogether]
 [-KeepWithNextParagraph] [-ContinueFormatting] [[-Supress] <Boolean>] [<CommonParameters>]
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
Position: 35
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AutoFit
{{Fill AutoFit Description}}

```yaml
Type: AutoFit
Parameter Sets: (All)
Aliases:
Accepted values: Contents, Window, ColumnWidth, Fixed

Required: False
Position: 4
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
Position: 16
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -BreakAcrossPages
{{Fill BreakAcrossPages Description}}

```yaml
Type: Boolean
Parameter Sets: (All)
Aliases:

Required: False
Position: 7
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -BreakPageAfterTable
{{Fill BreakPageAfterTable Description}}

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

### -BreakPageBeforeTable
{{Fill BreakPageBeforeTable Description}}

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

### -CapsStyle
{{Fill CapsStyle Description}}

```yaml
Type: CapsStyle[]
Parameter Sets: (All)
Aliases: CA
Accepted values: none, caps, smallCaps

Required: False
Position: 24
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
Position: 13
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ColummnWidth
{{Fill ColummnWidth Description}}

```yaml
Type: Single[]
Parameter Sets: (All)
Aliases:

Required: False
Position: 10
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ContinueFormatting
{{Fill ContinueFormatting Description}}

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

### -DataTable
{{Fill DataTable Description}}

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 3
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Design
{{Fill Design Description}}

```yaml
Type: TableDesign
Parameter Sets: (All)
Aliases:
Accepted values: Custom, TableNormal, TableGrid, LightShading, LightShadingAccent1, LightShadingAccent2, LightShadingAccent3, LightShadingAccent4, LightShadingAccent5, LightShadingAccent6, LightList, LightListAccent1, LightListAccent2, LightListAccent3, LightListAccent4, LightListAccent5, LightListAccent6, LightGrid, LightGridAccent1, LightGridAccent2, LightGridAccent3, LightGridAccent4, LightGridAccent5, LightGridAccent6, MediumShading1, MediumShading1Accent1, MediumShading1Accent2, MediumShading1Accent3, MediumShading1Accent4, MediumShading1Accent5, MediumShading1Accent6, MediumShading2, MediumShading2Accent1, MediumShading2Accent2, MediumShading2Accent3, MediumShading2Accent4, MediumShading2Accent5, MediumShading2Accent6, MediumList1, MediumList1Accent1, MediumList1Accent2, MediumList1Accent3, MediumList1Accent4, MediumList1Accent5, MediumList1Accent6, MediumList2, MediumList2Accent1, MediumList2Accent2, MediumList2Accent3, MediumList2Accent4, MediumList2Accent5, MediumList2Accent6, MediumGrid1, MediumGrid1Accent1, MediumGrid1Accent2, MediumGrid1Accent3, MediumGrid1Accent4, MediumGrid1Accent5, MediumGrid1Accent6, MediumGrid2, MediumGrid2Accent1, MediumGrid2Accent2, MediumGrid2Accent3, MediumGrid2Accent4, MediumGrid2Accent5, MediumGrid2Accent6, MediumGrid3, MediumGrid3Accent1, MediumGrid3Accent2, MediumGrid3Accent3, MediumGrid3Accent4, MediumGrid3Accent5, MediumGrid3Accent6, DarkList, DarkListAccent1, DarkListAccent2, DarkListAccent3, DarkListAccent4, DarkListAccent5, DarkListAccent6, ColorfulShading, ColorfulShadingAccent1, ColorfulShadingAccent2, ColorfulShadingAccent3, ColorfulShadingAccent4, ColorfulShadingAccent5, ColorfulShadingAccent6, ColorfulList, ColorfulListAccent1, ColorfulListAccent2, ColorfulListAccent3, ColorfulListAccent4, ColorfulListAccent5, ColorfulListAccent6, ColorfulGrid, ColorfulGridAccent1, ColorfulGridAccent2, ColorfulGridAccent3, ColorfulGridAccent4, ColorfulGridAccent5, ColorfulGridAccent6, None

Required: False
Position: 5
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Direction
{{Fill Direction Description}}

```yaml
Type: Direction
Parameter Sets: (All)
Aliases:
Accepted values: LeftToRight, RightToLeft

Required: False
Position: 6
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DirectionFormatting
{{Fill DirectionFormatting Description}}

```yaml
Type: Direction[]
Parameter Sets: (All)
Aliases:
Accepted values: LeftToRight, RightToLeft

Required: False
Position: 36
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DoNotAddTitle
{{Fill DoNotAddTitle Description}}

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

### -FontFamily
{{Fill FontFamily Description}}

```yaml
Type: String[]
Parameter Sets: (All)
Aliases: FontName

Required: False
Position: 15
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
Position: 14
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
Position: 26
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
Position: 31
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
Position: 23
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
Position: 33
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
Position: 34
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
Position: 17
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
Position: 30
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
Position: 29
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumColumns
{{Fill MaximumColumns Description}}

```yaml
Type: Int32
Parameter Sets: (All)
Aliases:

Required: False
Position: 8
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
Position: 28
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
Position: 39
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Paragraph
{{Fill Paragraph Description}}

```yaml
Type: InsertBeforeOrAfter
Parameter Sets: (All)
Aliases:

Required: False
Position: 1
Default value: None
Accept pipeline input: True (ByPropertyName, ByValue)
Accept wildcard characters: False
```

### -Percentage
{{Fill Percentage Description}}

```yaml
Type: Boolean
Parameter Sets: (All)
Aliases:

Required: False
Position: 12
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PercentageScale
{{Fill PercentageScale Description}}

```yaml
Type: Int32[]
Parameter Sets: (All)
Aliases:

Required: False
Position: 27
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
Position: 32
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
Position: 38
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
Position: 37
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
Position: 22
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
Position: 20
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
Position: 21
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
Position: 25
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
Position: 40
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Table
{{Fill Table Description}}

```yaml
Type: InsertBeforeOrAfter
Parameter Sets: (All)
Aliases:

Required: False
Position: 2
Default value: None
Accept pipeline input: True (ByPropertyName, ByValue)
Accept wildcard characters: False
```

### -TableWidth
{{Fill TableWidth Description}}

```yaml
Type: Single
Parameter Sets: (All)
Aliases:

Required: False
Position: 11
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Titles
{{Fill Titles Description}}

```yaml
Type: String[]
Parameter Sets: (All)
Aliases:

Required: False
Position: 9
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
Position: 19
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
Position: 18
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -WordDocument
{{Fill WordDocument Description}}

```yaml
Type: Container
Parameter Sets: (All)
Aliases:

Required: False
Position: 0
Default value: None
Accept pipeline input: True (ByPropertyName, ByValue)
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see about_CommonParameters (http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

### Xceed.Words.NET.Container
Xceed.Words.NET.InsertBeforeOrAfter

## OUTPUTS

### System.Object

## NOTES

## RELATED LINKS
