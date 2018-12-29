---
external help file: PSWriteWord-help.xml
Module Name: PSWriteWord
online version:
schema: 2.0.0
---

# Add-WordPageCount

## SYNOPSIS
{{Fill in the Synopsis}}

## SYNTAX

```
Add-WordPageCount [[-PageNumberFormat] <PageNumberFormat>] [[-Paragraph] <InsertBeforeOrAfter>]
 [[-Footer] <Footers>] [[-Header] <Headers>] [[-Alignment] <Alignment>] [[-Type] <String>] [[-Option] <String>]
 [[-TextBefore] <String>] [[-TextMiddle] <String>] [[-TextAfter] <String>] [[-Supress] <Boolean>]
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

### -Alignment
{{Fill Alignment Description}}

```yaml
Type: Alignment
Parameter Sets: (All)
Aliases:
Accepted values: left, center, right, both

Required: False
Position: 4
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Footer
{{Fill Footer Description}}

```yaml
Type: Footers
Parameter Sets: (All)
Aliases:

Required: False
Position: 2
Default value: None
Accept pipeline input: True (ByPropertyName, ByValue)
Accept wildcard characters: False
```

### -Header
{{Fill Header Description}}

```yaml
Type: Headers
Parameter Sets: (All)
Aliases:

Required: False
Position: 3
Default value: None
Accept pipeline input: True (ByPropertyName, ByValue)
Accept wildcard characters: False
```

### -Option
{{Fill Option Description}}

```yaml
Type: String
Parameter Sets: (All)
Aliases:
Accepted values: Both, PageCountOnly, PageNumberOnly

Required: False
Position: 6
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PageNumberFormat
{{Fill PageNumberFormat Description}}

```yaml
Type: PageNumberFormat
Parameter Sets: (All)
Aliases:
Accepted values: normal, roman

Required: False
Position: 0
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

### -Supress
{{Fill Supress Description}}

```yaml
Type: Boolean
Parameter Sets: (All)
Aliases:

Required: False
Position: 10
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TextAfter
{{Fill TextAfter Description}}

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 9
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TextBefore
{{Fill TextBefore Description}}

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 7
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TextMiddle
{{Fill TextMiddle Description}}

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 8
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Type
{{Fill Type Description}}

```yaml
Type: String
Parameter Sets: (All)
Aliases:
Accepted values: All, First, Even, Odd

Required: False
Position: 5
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see about_CommonParameters (http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

### Xceed.Words.NET.InsertBeforeOrAfter

### Xceed.Words.NET.Footers

### Xceed.Words.NET.Headers

## OUTPUTS

### System.Object
## NOTES

## RELATED LINKS
