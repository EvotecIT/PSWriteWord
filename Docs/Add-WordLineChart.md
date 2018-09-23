---
external help file: PSWriteWord-help.xml
Module Name: PSWriteWord
online version:
schema: 2.0.0
---

# Add-WordLineChart

## SYNOPSIS
{{Fill in the Synopsis}}

## SYNTAX

```
Add-WordLineChart [[-WordDocument] <Container>] [[-Paragraph] <InsertBeforeOrAfter>] [[-ChartName] <String>]
 [[-Names] <String[]>] [[-Values] <Int32[]>] [[-ChartSeries] <Series[]>]
 [[-ChartLegendPosition] <ChartLegendPosition>] [[-ChartLegendOverlay] <Boolean>] [-NoLegend]
 [<CommonParameters>]
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

### -ChartLegendOverlay
{{Fill ChartLegendOverlay Description}}

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

### -ChartLegendPosition
{{Fill ChartLegendPosition Description}}

```yaml
Type: ChartLegendPosition
Parameter Sets: (All)
Aliases:
Accepted values: Top, Bottom, Left, Right, TopRight

Required: False
Position: 6
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ChartName
{{Fill ChartName Description}}

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

### -ChartSeries
{{Fill ChartSeries Description}}

```yaml
Type: Series[]
Parameter Sets: (All)
Aliases:

Required: False
Position: 5
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Names
{{Fill Names Description}}

```yaml
Type: String[]
Parameter Sets: (All)
Aliases:

Required: False
Position: 3
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoLegend
{{Fill NoLegend Description}}

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

### -Values
{{Fill Values Description}}

```yaml
Type: Int32[]
Parameter Sets: (All)
Aliases:

Required: False
Position: 4
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
