---
external help file: PSWriteWord-help.xml
Module Name: PSWriteWord
online version:
schema: 2.0.0
---

# New-WordBlock

## SYNOPSIS
{{Fill in the Synopsis}}

## SYNTAX

```
New-WordBlock [-WordDocument] <Container> [[-TocGlobalDefinition] <Boolean>] [[-TocGlobalTitle] <String>]
 [[-TocGlobalRightTabPos] <Int32>] [[-TocGlobalSwitches] <TableOfContentsSwitches[]>] [[-TocEnable] <Boolean>]
 [[-TocText] <String>] [[-TocListLevel] <Int32>] [[-TocListItemType] <ListItemType>]
 [[-TocHeadingType] <HeadingType>] [[-EmptyParagraphsBefore] <Int32>] [[-EmptyParagraphsAfter] <Int32>]
 [[-PageBreaksBefore] <Int32>] [[-PageBreaksAfter] <Int32>] [[-Text] <String>] [[-TextNoData] <String>]
 [[-TextAlignment] <Nullable`1[]>] [[-TableData] <Object>] [[-TableDesign] <TableDesign>]
 [[-TableMaximumColumns] <Int32>] [[-TableTitleMerge] <Boolean>] [[-TableTitleText] <String>]
 [[-TableTitleAlignment] <Alignment>] [[-TableTitleColor] <Color>] [-TableTranspose]
 [[-TableColumnWidths] <Single[]>] [[-ListData] <Object>] [[-ListType] <ListItemType>]
 [[-ListTextEmpty] <String>] [[-ListBuilderContent] <String[]>] [[-ListBuilderType] <ListItemType[]>]
 [[-ListBuilderLevel] <Int32[]>] [[-TextBasedData] <Object>] [[-TextBasedDataAlignment] <Nullable`1[]>]
 [[-ChartEnable] <Boolean>] [[-ChartTitle] <String>] [[-ChartKeys] <Object>] [[-ChartValues] <Object>]
 [[-ChartLegendPosition] <ChartLegendPosition>] [[-ChartLegendOverlay] <Boolean>] [<CommonParameters>]
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

### -ChartEnable
{{Fill ChartEnable Description}}

```yaml
Type: Boolean
Parameter Sets: (All)
Aliases:

Required: False
Position: 33
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ChartKeys
{{Fill ChartKeys Description}}

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 35
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ChartLegendOverlay
{{Fill ChartLegendOverlay Description}}

```yaml
Type: Boolean
Parameter Sets: (All)
Aliases:

Required: False
Position: 38
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ChartLegendPosition
{{Fill ChartLegendPosition Description}}

```yaml
Type: ChartLegendPosition
Parameter Sets: (All)
Aliases:
Accepted values: Top, Bottom, Left, Right, TopRight

Required: False
Position: 37
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ChartTitle
{{Fill ChartTitle Description}}

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 34
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ChartValues
{{Fill ChartValues Description}}

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 36
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -EmptyParagraphsAfter
{{Fill EmptyParagraphsAfter Description}}

```yaml
Type: Int32
Parameter Sets: (All)
Aliases:

Required: False
Position: 11
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -EmptyParagraphsBefore
{{Fill EmptyParagraphsBefore Description}}

```yaml
Type: Int32
Parameter Sets: (All)
Aliases:

Required: False
Position: 10
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ListBuilderContent
{{Fill ListBuilderContent Description}}

```yaml
Type: String[]
Parameter Sets: (All)
Aliases:

Required: False
Position: 28
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ListBuilderLevel
{{Fill ListBuilderLevel Description}}

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

### -ListBuilderType
{{Fill ListBuilderType Description}}

```yaml
Type: ListItemType[]
Parameter Sets: (All)
Aliases:
Accepted values: Bulleted, Numbered

Required: False
Position: 29
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ListData
{{Fill ListData Description}}

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 25
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ListTextEmpty
{{Fill ListTextEmpty Description}}

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 27
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ListType
{{Fill ListType Description}}

```yaml
Type: ListItemType
Parameter Sets: (All)
Aliases:
Accepted values: Bulleted, Numbered

Required: False
Position: 26
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PageBreaksAfter
{{Fill PageBreaksAfter Description}}

```yaml
Type: Int32
Parameter Sets: (All)
Aliases:

Required: False
Position: 13
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PageBreaksBefore
{{Fill PageBreaksBefore Description}}

```yaml
Type: Int32
Parameter Sets: (All)
Aliases:

Required: False
Position: 12
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TableColumnWidths
{{Fill TableColumnWidths Description}}

```yaml
Type: Single[]
Parameter Sets: (All)
Aliases:

Required: False
Position: 24
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TableData
{{Fill TableData Description}}

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 17
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TableDesign
{{Fill TableDesign Description}}

```yaml
Type: TableDesign
Parameter Sets: (All)
Aliases:
Accepted values: Custom, TableNormal, TableGrid, LightShading, LightShadingAccent1, LightShadingAccent2, LightShadingAccent3, LightShadingAccent4, LightShadingAccent5, LightShadingAccent6, LightList, LightListAccent1, LightListAccent2, LightListAccent3, LightListAccent4, LightListAccent5, LightListAccent6, LightGrid, LightGridAccent1, LightGridAccent2, LightGridAccent3, LightGridAccent4, LightGridAccent5, LightGridAccent6, MediumShading1, MediumShading1Accent1, MediumShading1Accent2, MediumShading1Accent3, MediumShading1Accent4, MediumShading1Accent5, MediumShading1Accent6, MediumShading2, MediumShading2Accent1, MediumShading2Accent2, MediumShading2Accent3, MediumShading2Accent4, MediumShading2Accent5, MediumShading2Accent6, MediumList1, MediumList1Accent1, MediumList1Accent2, MediumList1Accent3, MediumList1Accent4, MediumList1Accent5, MediumList1Accent6, MediumList2, MediumList2Accent1, MediumList2Accent2, MediumList2Accent3, MediumList2Accent4, MediumList2Accent5, MediumList2Accent6, MediumGrid1, MediumGrid1Accent1, MediumGrid1Accent2, MediumGrid1Accent3, MediumGrid1Accent4, MediumGrid1Accent5, MediumGrid1Accent6, MediumGrid2, MediumGrid2Accent1, MediumGrid2Accent2, MediumGrid2Accent3, MediumGrid2Accent4, MediumGrid2Accent5, MediumGrid2Accent6, MediumGrid3, MediumGrid3Accent1, MediumGrid3Accent2, MediumGrid3Accent3, MediumGrid3Accent4, MediumGrid3Accent5, MediumGrid3Accent6, DarkList, DarkListAccent1, DarkListAccent2, DarkListAccent3, DarkListAccent4, DarkListAccent5, DarkListAccent6, ColorfulShading, ColorfulShadingAccent1, ColorfulShadingAccent2, ColorfulShadingAccent3, ColorfulShadingAccent4, ColorfulShadingAccent5, ColorfulShadingAccent6, ColorfulList, ColorfulListAccent1, ColorfulListAccent2, ColorfulListAccent3, ColorfulListAccent4, ColorfulListAccent5, ColorfulListAccent6, ColorfulGrid, ColorfulGridAccent1, ColorfulGridAccent2, ColorfulGridAccent3, ColorfulGridAccent4, ColorfulGridAccent5, ColorfulGridAccent6, None

Required: False
Position: 18
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TableMaximumColumns
{{Fill TableMaximumColumns Description}}

```yaml
Type: Int32
Parameter Sets: (All)
Aliases:

Required: False
Position: 19
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TableTitleAlignment
{{Fill TableTitleAlignment Description}}

```yaml
Type: Alignment
Parameter Sets: (All)
Aliases:
Accepted values: left, center, right, both

Required: False
Position: 22
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TableTitleColor
{{Fill TableTitleColor Description}}

```yaml
Type: Color
Parameter Sets: (All)
Aliases:

Required: False
Position: 23
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TableTitleMerge
{{Fill TableTitleMerge Description}}

```yaml
Type: Boolean
Parameter Sets: (All)
Aliases:

Required: False
Position: 20
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TableTitleText
{{Fill TableTitleText Description}}

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 21
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TableTranspose
{{Fill TableTranspose Description}}

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

### -Text
{{Fill Text Description}}

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 14
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TextAlignment
{{Fill TextAlignment Description}}

```yaml
Type: Nullable`1[]
Parameter Sets: (All)
Aliases:

Required: False
Position: 16
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TextBasedData
{{Fill TextBasedData Description}}

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 31
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TextBasedDataAlignment
{{Fill TextBasedDataAlignment Description}}

```yaml
Type: Nullable`1[]
Parameter Sets: (All)
Aliases:

Required: False
Position: 32
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TextNoData
{{Fill TextNoData Description}}

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 15
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TocEnable
{{Fill TocEnable Description}}

```yaml
Type: Boolean
Parameter Sets: (All)
Aliases:

Required: False
Position: 5
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TocGlobalDefinition
{{Fill TocGlobalDefinition Description}}

```yaml
Type: Boolean
Parameter Sets: (All)
Aliases:

Required: False
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TocGlobalRightTabPos
{{Fill TocGlobalRightTabPos Description}}

```yaml
Type: Int32
Parameter Sets: (All)
Aliases:

Required: False
Position: 3
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TocGlobalSwitches
{{Fill TocGlobalSwitches Description}}

```yaml
Type: TableOfContentsSwitches[]
Parameter Sets: (All)
Aliases:
Accepted values: None, A, B, C, D, F, H, L, N, O, P, S, T, U, W, X, Z

Required: False
Position: 4
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TocGlobalTitle
{{Fill TocGlobalTitle Description}}

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

### -TocHeadingType
{{Fill TocHeadingType Description}}

```yaml
Type: HeadingType
Parameter Sets: (All)
Aliases:
Accepted values: Heading1, Heading2, Heading3, Heading4, Heading5, Heading6, Heading7, Heading8, Heading9

Required: False
Position: 9
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TocListItemType
{{Fill TocListItemType Description}}

```yaml
Type: ListItemType
Parameter Sets: (All)
Aliases:
Accepted values: Bulleted, Numbered

Required: False
Position: 8
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TocListLevel
{{Fill TocListLevel Description}}

```yaml
Type: Int32
Parameter Sets: (All)
Aliases:

Required: False
Position: 7
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TocText
{{Fill TocText Description}}

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 6
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

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByPropertyName, ByValue)
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see about_CommonParameters (http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

### System.Object

## OUTPUTS

### System.Object

## NOTES

## RELATED LINKS
