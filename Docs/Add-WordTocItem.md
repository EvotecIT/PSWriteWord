---
external help file: PSWriteWord-help.xml
Module Name: PSWriteWord
online version:
schema: 2.0.0
---

# Add-WordTocItem

## SYNOPSIS
{{Fill in the Synopsis}}

## SYNTAX

```
Add-WordTocItem [[-WordDocument] <Container>] [[-ListLevel] <Int32>] [[-ListItemType] <ListItemType>]
 [[-Text] <Object>] [[-HeadingType] <HeadingType>] [[-StartNumber] <Int32>] [[-TrackChanges] <Boolean>]
 [[-ContinueNumbering] <Boolean>] [[-Supress] <Boolean>] [<CommonParameters>]
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

### -ContinueNumbering
{{Fill ContinueNumbering Description}}

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

### -HeadingType
{{Fill HeadingType Description}}

```yaml
Type: HeadingType
Parameter Sets: (All)
Aliases: HT
Accepted values: Heading1, Heading2, Heading3, Heading4, Heading5, Heading6, Heading7, Heading8, Heading9

Required: False
Position: 4
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ListItemType
{{Fill ListItemType Description}}

```yaml
Type: ListItemType
Parameter Sets: (All)
Aliases: ListType
Accepted values: Bulleted, Numbered

Required: False
Position: 2
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ListLevel
{{Fill ListLevel Description}}

```yaml
Type: Int32
Parameter Sets: (All)
Aliases: Level

Required: False
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -StartNumber
{{Fill StartNumber Description}}

```yaml
Type: Int32
Parameter Sets: (All)
Aliases:

Required: False
Position: 5
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
Position: 8
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Text
{{Fill Text Description}}

```yaml
Type: Object
Parameter Sets: (All)
Aliases: Value, ListValue

Required: False
Position: 3
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TrackChanges
{{Fill TrackChanges Description}}

```yaml
Type: Boolean
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

Required: False
Position: 0
Default value: None
Accept pipeline input: True (ByPropertyName, ByValue)
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable.
For more information, see about_CommonParameters (http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

### Xceed.Words.NET.Container


## OUTPUTS

### System.Object

## NOTES

## RELATED LINKS
