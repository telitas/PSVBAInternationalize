# PSVBAInternationalize

VBA source files internationalize tool.

## How to Install

You can install from [PowerShell Gallery](https://www.powershellgallery.com/packages/PSVBAInternationalize/).

```ps1
Install-Module -Name PSVBAInternationalize -Scope CurrentUser
```

## Usage

Assume you have the following vba source code.
The string literal enclosed in "_(" and ")" are regarded as placeholder.

```vb
' HelloWorldBase.bas
Public Sub HelloWorld()
Attribute HelloWorld.VB_Description = "_(Description_HelloWorld)"
    Call MsgBox("_(HelloWorld)")
End Sub
```

First, export placeholders from source code.

```ps1
Export-VBATranslationPlaceHolder -SourcePath .\HelloWorldBase.bas -DestinationPath HelloWorldTranslation.xml
```

and you will get the folloing file.

```xml
<!-- HelloWorldTranslation.xml --> 
<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<translations>
<translation id="Description_HelloWorld"></translation>
<translation id="HelloWorld"></translation>
</translations>
```

Next, fill the translation file

```xml
<!-- HelloWorldTranslation.xml --> 
<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<translations>
<translation id="Description_HelloWorld">This is hello, world subroutine.</translation>
<translation id="HelloWorld">hello, world</translation>
</translations>
```

and resolve placeholders.

```ps1
Resolve-VBATranslationPlaceHolder -SourcePath .\HelloWorldBase.bas -TranslationPath HelloWorldTranslation.xml -DestinationPath .\HelloWorld.bas
```

Finally, you will get the translated source code.

```vb
' HelloWorld.bas
Public Sub HelloWorld()
Attribute HelloWorld.VB_Description = "This is hello, world subroutine."
    Call MsgBox("hello, world")
End Sub
```

## NOTE

This module imprementation is very lazy because I wish VBA would be replaced
to other languages in the near future.

## License

MIT

Copyright (c) 2022 telitas

See the LICENSE file or https://opensource.org/licenses/mit-license.php for details.
