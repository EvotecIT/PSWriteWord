---
external help file: PSWriteWord-help.xml
Module Name: PSWriteWord
online version:
schema: 2.0.0
---

# Add-WordPicture

## SYNOPSIS
{{Fill in the Synopsis}}

## SYNTAX

```
Add-WordPicture [[-WordDocument] <Container>] [[-Paragraph] <InsertBeforeOrAfter>] [[-Picture] <Picture>]
 [[-ImagePath] <String>] [[-Alignment] <Alignment>] [[-Rotation] <Int32>] [-FlipHorizontal] [-FlipVertical]
 [[-ImageWidth] <Int32>] [[-ImageHeight] <Int32>] [[-Description] <String>] [[-Supress] <Boolean>]
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

### -Description
{{Fill Description Description}}

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

### -FlipHorizontal
{{Fill FlipHorizontal Description}}

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

### -FlipVertical
{{Fill FlipVertical Description}}

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

### -ImageHeight
{{Fill ImageHeight Description}}

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

### -ImagePath
{{Fill ImagePath Description}}

```yaml
Type: String
Parameter Sets: (All)
Aliases: FileImagePath

Required: False
Position: 3
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ImageWidth
{{Fill ImageWidth Description}}

```yaml
Type: Int32
Parameter Sets: (All)
Aliases:

Required: False
Position: 6
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

### -Picture
{{Fill Picture Description}}

```yaml
Type: Picture
Parameter Sets: (All)
Aliases:

Required: False
Position: 2
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Rotation
{{Fill Rotation Description}}

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
Position: 9
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
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

### Container
InsertBeforeOrAfter

## OUTPUTS

### System.Object

## NOTES

## RELATED LINKS
