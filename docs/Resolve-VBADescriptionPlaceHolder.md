---
external help file: PSVBAInternationalize-help.xml
Module Name: PSVBAInternationalize
online version:
schema: 2.0.0
---

# Resolve-VBATranslationPlaceHolder

## SYNOPSIS
Resolve placeholders for translation in VBA source code.

## SYNTAX

```
Resolve-VBATranslationPlaceHolder [-SourcePath] <String> [[-SourceEncoding] <Encoding>]
 [-TranslationPath] <String> [-DestinationPath] <String> [[-DestinationEncoding] <Encoding>] [-Force]
 [<CommonParameters>]
```

## DESCRIPTION
Resolve placeholders for translation in VBA source code.
The translation file must be xml file of the prescribed format.

## EXAMPLES

### Example 1
```powershell
PS C:\> Resolve-VBATranslationPlaceHolder -SourcePath .\path\to\source.bas -TranslationPath .\path\to\translation.xml -DestinationPath .\path\to\destination.bas
```

Resolve placeholders in ".\path\to\source.bas" with translate definitions in ".\path\to\translation.xml".
The resolved file will saved at ".\path\to\destination.bas"

## PARAMETERS

### -DestinationEncoding
Encoding of the destination file.

```yaml
Type: Encoding
Parameter Sets: (All)
Aliases:

Required: False
Position: 4
Default value: System.Text.UTF8Encoding
Accept pipeline input: False
Accept wildcard characters: False
```

### -DestinationPath
The location where the translated VBA file will be saved.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: True
Position: 3
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SourceEncoding
Encoding of the source file.

```yaml
Type: Encoding
Parameter Sets: (All)
Aliases:

Required: False
Position: 1
Default value: System.Text.UTF8Encoding
Accept pipeline input: False
Accept wildcard characters: False
```

### -SourcePath
A VBA source file where the placeholder will be replaced.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TranslationPath
An xml file containing the translations.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: True
Position: 2
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Force
When the output file already exist, it will be overwritten.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

### None
## OUTPUTS

### System.Object
## NOTES

## RELATED LINKS
