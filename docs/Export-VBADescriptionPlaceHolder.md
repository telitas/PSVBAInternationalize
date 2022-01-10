---
external help file: PSVBAInternationalize-help.xml
Module Name: PSVBAInternationalize
online version:
schema: 2.0.0
---

# Export-VBATranslationPlaceHolder

## SYNOPSIS
Export placeholders for translation from VBA source code.

## SYNTAX

```
Export-VBATranslationPlaceHolder [-SourcePath] <String> [[-SourceEncoding] <Encoding>]
 [-DestinationPath] <String> [-Force] [<CommonParameters>]
```

## DESCRIPTION
Export placeholders for translation from VBA source code.
The exported placeholders will be converted into an xml file of the prescribed format.

## EXAMPLES

### Example 1
```powershell
PS C:\> Export-VBATranslationPlaceHolder -SourcePath .\path\to\source.bas -DestinationPath .\path\to\translation.xml
```

Export placeholders from ".\path\to\source.bas".
The output XML file is saved at ".\path\to\translation.xml"

### Example 2
```powershell
PS C:\> Export-VBATranslationPlaceHolder -SourcePath .\path\to\source.bas -SourceEncoding [System.Text.Encoding]::Unicode -DestinationPath .\path\to\translation.xml -DestinationPath .\path\to\translation.xml -DestinationEncoding [System.Text.Encoding]::UTF32
```

Export placeholders from ".\path\to\source.bas" with utf-16 encoding.
The output XML file is saved at ".\path\to\translation.xml" with utf-32 encoding.

## PARAMETERS

### -DestinationPath
The location where the xml file that containing exported placeholders will be saved.

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
A VBA source file which you want to export placeholders.

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

### -Force
When output file already exist, it will be overwritten.

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
