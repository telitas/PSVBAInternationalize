Describe 'PSVBAInternationalize' {
    BeforeAll {
        function WrappedOutFile{
            Param(
                [Parameter(ValueFromPipeline)][string]$InputObject = "",
                [string]$Path,
                [System.Text.Encoding]$Encoding = (New-Object -TypeName System.Text.UTF8Encoding),
                [switch]$NoNewLine
            )
            process {
                $writer = New-Object -TypeName System.IO.StreamWriter -ArgumentList @($Path, $false, $Encoding)
                $writer.NewLine = "`r`n"
                try
                {
                    $writer.Write(($InputObject -split "`r?`n") -join $writer.NewLine)
                    if(-not $NoNewLine)
                    {
                        $writer.Write($writer.NewLine)
                    }
                }
                finally
                {
                    $writer.Close()
                }
            }
        }
        function WrappedGetContent{
            Param(
                [string]$Path,
                [System.Text.Encoding]$Encoding = (New-Object -TypeName System.Text.UTF8Encoding),
                [switch]$Raw
            )
            process {
                $reader = New-Object -TypeName System.IO.StreamReader -ArgumentList @($Path, $Encoding)
                try
                {
                    $content =$reader.ReadToEnd()
                    if($content -eq "")
                    {
                        $null
                    }
                    else
                    {
                        if ($Raw) {
                            $content
                        }
                        else
                        {
                            $content -split "`r?`n"
                        }
                    }
                }
                finally
                {
                    $reader.Close()
                }
            }
        }
        function ReplaceLFToCRLF{
            param(
                [string]$Content
            )
            $Content  -replace "`n(?<!`r`n)", "`r`n"
        }
        Set-StrictMode -Version Latest
        $Testee = Join-Path -Path $PSScriptRoot -ChildPath ((Split-Path -Leaf $PSCommandPath) -replace '\.Tests\.ps1$', '.psm1')
        Import-Module $Testee -Force -ErrorAction Stop
        $WorkDirectory = Join-Path -Path $PSScriptRoot -ChildPath ".work"
        if(Test-Path -Path $WorkDirectory)
        {
            Write-Error "prease remove work directory `"$($WorkDirectory)`" manually." -ErrorAction Stop
        }
        New-Item -Path $WorkDirectory -ItemType Directory
    }
    AfterAll {
        Remove-Module 'PSVBAInternationalize'
        Remove-Item -Path $WorkDirectory -Recurse -Force
    }
    Context 'Lint' {
        It 'lint' {
            Get-ChildItem -Path .\src\*.ps1 -Recurse | ForEach-Object -Process {
                $output = Invoke-ScriptAnalyzer $_
                $output | Should -BeNullOrEmpty
            }
        }
    }
    Context 'Export-VBATranslationPlaceHolder' {
        Context 'Passed' {
            It 'Empty file' {
                $inputFile = Join-Path -Path $WorkDirectory -ChildPath "in.bas"
                $outputFile = Join-Path -Path $WorkDirectory -ChildPath "out.xml"
                $encoding = (New-Object -TypeName System.Text.UTF8Encoding -ArgumentList @($false))

                @"
"@ | WrappedOutFile -Path $inputFile -Encoding $encoding -NoNewline
                Export-VBATranslationPlaceHolder -SourcePath $inputfile -SourceEncoding $encoding -DestinationPath $outputFile
                WrappedGetContent -Path $outputFile -Raw | Should -Be @"
<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<translations />
"@

                Remove-Item $inputFile -Force
                Remove-Item $outputFile -Force
            }
            It 'No placeholder' {
                $inputFile = Join-Path -Path $WorkDirectory -ChildPath "in.bas"
                $outputFile = Join-Path -Path $WorkDirectory -ChildPath "out.xml"
                $encoding = (New-Object -TypeName System.Text.UTF8Encoding -ArgumentList @($false))

                @"
Attribute VB_Description = "This is dummy class."
Public DummyField As String
Attribute DummyField.VB_VarDescription = "This is dummy field."
Public Sub DummyMethod()
Attribute DummyMethod.VB_Description = "This is dummy method."
    Debug.Print "This is dummy string 1"
    Debug.Print "This is dummy string 2" & "This is dummy string 3"
    Debug.Print "This is dummy string 4": Debug.Print "This is dummy string 5"
    Debug.Print """_(NotPlaceHolder)"""
End Sub
"@ | WrappedOutFile -Path $inputFile -Encoding $encoding -NoNewline
                Export-VBATranslationPlaceHolder -SourcePath $inputfile -SourceEncoding $encoding -DestinationPath $outputFile
                WrappedGetContent -Path $outputFile -Raw | Should -Be @"
<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<translations />
"@

                Remove-Item $inputFile -Force
                Remove-Item $outputFile -Force
            }
            It '1 normal placeholder' {
                $inputFile = Join-Path -Path $WorkDirectory -ChildPath "in.bas"
                $outputFile = Join-Path -Path $WorkDirectory -ChildPath "out.xml"
                $encoding = (New-Object -TypeName System.Text.UTF8Encoding -ArgumentList @($false))

                @"
Public DummyField As String
Public Sub DummyMethod()
    Debug.Print "_(NormalPlaceHolder_1)"
End Sub
"@ | WrappedOutFile -Path $inputFile -Encoding $encoding -NoNewline
                Export-VBATranslationPlaceHolder -SourcePath $inputfile -SourceEncoding $encoding -DestinationPath $outputFile
                WrappedGetContent -Path $outputFile -Raw | Should -Be @"
<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<translations>
<translation id="NormalPlaceHolder_1"></translation>
</translations>
"@
                Remove-Item $inputFile -Force
                Remove-Item $outputFile -Force
            }
            It 'More than 1 placeholders' {
                $inputFile = Join-Path -Path $WorkDirectory -ChildPath "in.bas"
                $outputFile = Join-Path -Path $WorkDirectory -ChildPath "out.xml"
                $encoding = (New-Object -TypeName System.Text.UTF8Encoding -ArgumentList @($false))

                @"
Public DummyField As String
Public Sub DummyMethod()
    Debug.Print "_(NormalPlaceHolder_1)"
    Debug.Print "_(NormalPlaceHolder_2)"
End Sub
"@ | WrappedOutFile -Path $inputFile -Encoding $encoding -NoNewline
                Export-VBATranslationPlaceHolder -SourcePath $inputfile -SourceEncoding $encoding -DestinationPath $outputFile
                WrappedGetContent -Path $outputFile -Raw | Should -Be @"
<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<translations>
<translation id="NormalPlaceHolder_1"></translation>
<translation id="NormalPlaceHolder_2"></translation>
</translations>
"@

                Remove-Item $inputFile -Force
                Remove-Item $outputFile -Force
            }
            It 'More than 1 placeholders inline' {
                $inputFile = Join-Path -Path $WorkDirectory -ChildPath "in.bas"
                $outputFile = Join-Path -Path $WorkDirectory -ChildPath "out.xml"
                $encoding = (New-Object -TypeName System.Text.UTF8Encoding -ArgumentList @($false))

                @"
Public Sub DummyMethod()
    Debug.Print "_(NormalPlaceHolder_1)" & "_(NormalPlaceHolder_2)"
    Debug.Print "_(NormalPlaceHolder_3)": Debug.Print "_(NormalPlaceHolder_4)"
End Sub
"@ | WrappedOutFile -Path $inputFile -Encoding $encoding -NoNewline
                Export-VBATranslationPlaceHolder -SourcePath $inputfile -SourceEncoding $encoding -DestinationPath $outputFile
                WrappedGetContent -Path $outputFile -Raw | Should -Be @"
<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<translations>
<translation id="NormalPlaceHolder_1"></translation>
<translation id="NormalPlaceHolder_2"></translation>
<translation id="NormalPlaceHolder_3"></translation>
<translation id="NormalPlaceHolder_4"></translation>
</translations>
"@

                Remove-Item $inputFile -Force
                Remove-Item $outputFile -Force
            }
            It 'Class description placeholders' {
                $inputFile = Join-Path -Path $WorkDirectory -ChildPath "in.bas"
                $outputFile = Join-Path -Path $WorkDirectory -ChildPath "out.xml"
                $encoding = (New-Object -TypeName System.Text.UTF8Encoding -ArgumentList @($false))

                @"
Public DummyField As String
Attribute VB_Description = "_(ClassDescriptionPlaceHolder)"
Public Sub DummyMethod()
End Sub
"@ | WrappedOutFile -Path $inputFile -Encoding $encoding -NoNewline
                Export-VBATranslationPlaceHolder -SourcePath $inputfile -SourceEncoding $encoding -DestinationPath $outputFile
                WrappedGetContent -Path $outputFile -Raw | Should -Be @"
<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<translations>
<translation id="ClassDescriptionPlaceHolder"></translation>
</translations>
"@

                Remove-Item $inputFile -Force
                Remove-Item $outputFile -Force
            }
            It 'Field description placeholders' {
                $inputFile = Join-Path -Path $WorkDirectory -ChildPath "in.bas"
                $outputFile = Join-Path -Path $WorkDirectory -ChildPath "out.xml"
                $encoding = (New-Object -TypeName System.Text.UTF8Encoding -ArgumentList @($false))

                @"
Public DummyField As String
Attribute DummyField.VB_VarDescription = "_(FieldDescriptionPlaceHolder)"
Public Sub DummyMethod()
End Sub
"@ | WrappedOutFile -Path $inputFile -Encoding $encoding -NoNewline
                Export-VBATranslationPlaceHolder -SourcePath $inputfile -SourceEncoding $encoding -DestinationPath $outputFile
                WrappedGetContent -Path $outputFile -Raw | Should -Be @"
<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<translations>
<translation id="FieldDescriptionPlaceHolder"></translation>
</translations>
"@

                Remove-Item $inputFile -Force
                Remove-Item $outputFile -Force
            }
            It 'Method description placeholders' {
                $inputFile = Join-Path -Path $WorkDirectory -ChildPath "in.bas"
                $outputFile = Join-Path -Path $WorkDirectory -ChildPath "out.xml"
                $encoding = (New-Object -TypeName System.Text.UTF8Encoding -ArgumentList @($false))

                @"
Public DummyField As String
Public Sub DummyMethod()
Attribute DummyMethod.VB_Description = "_(MethodDescriptionPlaceHolder)"
End Sub
"@ | WrappedOutFile -Path $inputFile -Encoding $encoding -NoNewline
                Export-VBATranslationPlaceHolder -SourcePath $inputfile -SourceEncoding $encoding -DestinationPath $outputFile
                WrappedGetContent -Path $outputFile -Raw | Should -Be @"
<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<translations>
<translation id="MethodDescriptionPlaceHolder"></translation>
</translations>
"@

                Remove-Item $inputFile -Force
                Remove-Item $outputFile -Force
            }
            It 'All pattern' {
                $inputFile = Join-Path -Path $WorkDirectory -ChildPath "in.bas"
                $outputFile = Join-Path -Path $WorkDirectory -ChildPath "out.xml"
                $encoding = (New-Object -TypeName System.Text.UTF8Encoding -ArgumentList @($false))

                @"
Attribute VB_Description = "_(ClassDescriptionPlaceHolder)"
Public DummyField As String
Attribute DummyField.VB_VarDescription = "_(FieldDescriptionPlaceHolder)"
Public Sub DummyMethod()
Attribute DummyMethod.VB_Description = "_(MethodDescriptionPlaceHolder)"
    Debug.Print "_(NormalPlaceHolder_1)"
    Debug.Print "_(NormalPlaceHolder_2)" & "_(NormalPlaceHolder_3)"
    Debug.Print "_(NormalPlaceHolder_4)": Debug.Print "_(NormalPlaceHolder_5)"
    Debug.Print """_(NotPlaceHolder)"""
End Sub

"@ | WrappedOutFile -Path $inputFile -Encoding $encoding -NoNewline
                Export-VBATranslationPlaceHolder -SourcePath $inputfile -SourceEncoding $encoding -DestinationPath $outputFile
                WrappedGetContent -Path $outputFile -Raw | Should -Be @"
<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<translations>
<translation id="ClassDescriptionPlaceHolder"></translation>
<translation id="FieldDescriptionPlaceHolder"></translation>
<translation id="MethodDescriptionPlaceHolder"></translation>
<translation id="NormalPlaceHolder_1"></translation>
<translation id="NormalPlaceHolder_2"></translation>
<translation id="NormalPlaceHolder_3"></translation>
<translation id="NormalPlaceHolder_4"></translation>
<translation id="NormalPlaceHolder_5"></translation>
</translations>
"@

                Remove-Item $inputFile -Force
                Remove-Item $outputFile -Force
            }
            It 'File already exists with force' {
                $inputFile = Join-Path -Path $WorkDirectory -ChildPath "in.bas"
                $outputFile = Join-Path -Path $WorkDirectory -ChildPath "out.xml"
                $encoding = (New-Object -TypeName System.Text.UTF8Encoding -ArgumentList @($false))

                @"
"@ | WrappedOutFile -Path $inputFile -Encoding $encoding -NoNewline
@"
"@ | WrappedOutFile -Path $outputFile -Encoding $encoding -NoNewline
                Export-VBATranslationPlaceHolder -SourcePath $inputfile -SourceEncoding $encoding -DestinationPath $outputFile -Force

                Remove-Item $inputFile -Force
                Remove-Item $outputFile -Force
            }
            It 'Create a parent directory if it does not exists.' {
                $inputFile = Join-Path -Path $WorkDirectory -ChildPath "in.bas"
                $outputFile = Join-Path -Path $WorkDirectory -ChildPath "out\out.xml"
                $outputParent = Split-Path -Path $outputFile -Parent
                $encoding = (New-Object -TypeName System.Text.UTF8Encoding -ArgumentList @($false))

                @"
"@ | WrappedOutFile -Path $inputFile -Encoding $encoding -NoNewline
                Export-VBATranslationPlaceHolder -SourcePath $inputfile -SourceEncoding $encoding -DestinationPath $outputFile
                Test-Path -Path $outputParent | Should -BeTrue

                Remove-Item $inputFile -Force
                Remove-Item $outputFile -Force
                Remove-Item $outputParent -Force
            }
        }
        Context 'Exception' {
            It 'File already exists without force' {
                $inputFile = Join-Path -Path $WorkDirectory -ChildPath "in.bas"
                $outputFile = Join-Path -Path $WorkDirectory -ChildPath "out.xml"
                $encoding = (New-Object -TypeName System.Text.UTF8Encoding -ArgumentList @($false))

                @"
"@ | WrappedOutFile -Path $inputFile -Encoding $encoding -NoNewline
                @"
"@ | WrappedOutFile -Path $outputFile -Encoding $encoding -NoNewline
                {Export-VBATranslationPlaceHolder -SourcePath $inputfile -SourceEncoding $encoding -DestinationPath $outputFile} | Should -Throw

                Remove-Item $inputFile -Force
                Remove-Item $outputFile -Force
            }
        }
    }
    Context 'Resolve-VBATranslationPlaceHolder' {
        Context 'Passed' {
            It 'Empty source file' {
                $inputFile = Join-Path -Path $WorkDirectory -ChildPath "in.bas"
                $translationFile = Join-Path -Path $WorkDirectory -ChildPath "out.xml"
                $outputFile = Join-Path -Path $WorkDirectory -ChildPath "out.bas"
                $encoding = (New-Object -TypeName System.Text.UTF8Encoding -ArgumentList @($false))

                @"
"@ | WrappedOutFile -Path $inputFile -Encoding $encoding -NoNewline
                @"
<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<translations>
<translation id="ClassDescriptionPlaceHolder">This is class description.</translation>
<translation id="FieldDescriptionPlaceHolder">This is field description.</translation>
<translation id="MethodDescriptionPlaceHolder">This is method description.</translation>
<translation id="NormalPlaceHolder_1">This is normal string 1.</translation>
<translation id="NormalPlaceHolder_2">This is normal string 2.</translation>
<translation id="NormalPlaceHolder_3">This is normal string 3.</translation>
<translation id="NormalPlaceHolder_4">This is normal string 4.</translation>
<translation id="NormalPlaceHolder_5">This is normal string 5.</translation>
</translations>
"@ | WrappedOutFile -Path $translationFile -Encoding $encoding -NoNewline
                Resolve-VBATranslationPlaceHolder -SourcePath $inputfile -SourceEncoding $encoding -TranslationPath $translationFile -DestinationPath $outputFile -DestinationEncoding $encoding
                WrappedGetContent -Path $outputFile -Raw | Should -Be $null

                Remove-Item $inputFile -Force
                Remove-Item $translationFile -Force
                Remove-Item $outputFile -Force
            }
            It 'Empty translate file' {
                $inputFile = Join-Path -Path $WorkDirectory -ChildPath "in.bas"
                $translationFile = Join-Path -Path $WorkDirectory -ChildPath "out.xml"
                $outputFile = Join-Path -Path $WorkDirectory -ChildPath "out.bas"
                $encoding = (New-Object -TypeName System.Text.UTF8Encoding -ArgumentList @($false))

                @"
Public DummyField As String
Public Sub DummyMethod()
End Sub
"@ | WrappedOutFile -Path $inputFile -Encoding $encoding -NoNewline
                @"
<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<translations />
"@ | WrappedOutFile -Path $translationFile -Encoding $encoding -NoNewline
                Resolve-VBATranslationPlaceHolder -SourcePath $inputfile -SourceEncoding $encoding -TranslationPath $translationFile -DestinationPath $outputFile -DestinationEncoding $encoding

                WrappedGetContent -Path $outputFile -Raw | Should -Be (ReplaceLFToCRLF -Content @"
Public DummyField As String
Public Sub DummyMethod()
End Sub
"@)
                Remove-Item $inputFile -Force
                Remove-Item $translationFile -Force
                Remove-Item $outputFile -Force
            }
            It '1 placeholder' {
                $inputFile = Join-Path -Path $WorkDirectory -ChildPath "in.bas"
                $translationFile = Join-Path -Path $WorkDirectory -ChildPath "out.xml"
                $outputFile = Join-Path -Path $WorkDirectory -ChildPath "out.bas"
                $encoding = (New-Object -TypeName System.Text.UTF8Encoding -ArgumentList @($false))

                @"
Public DummyField As String
Public Sub DummyMethod()
    Debug.Print "_(NormalPlaceHolder_1)"
End Sub
"@ | WrappedOutFile -Path $inputFile -Encoding $encoding -NoNewline
                @"
<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<translations>
<translation id="NormalPlaceHolder_1">This is normal string 1.</translation>
</translations>
"@ | WrappedOutFile -Path $translationFile -Encoding $encoding -NoNewline

                Resolve-VBATranslationPlaceHolder -SourcePath $inputfile -SourceEncoding $encoding -TranslationPath $translationFile -DestinationPath $outputFile -DestinationEncoding $encoding

                WrappedGetContent -Path $outputFile -Raw | Should -Be (ReplaceLFToCRLF -Content @"
Public DummyField As String
Public Sub DummyMethod()
    Debug.Print "This is normal string 1."
End Sub
"@)

                Remove-Item $inputFile -Force
                Remove-Item $translationFile -Force
                Remove-Item $outputFile -Force
            }
            It 'More than 1 placeholders' {
                $inputFile = Join-Path -Path $WorkDirectory -ChildPath "in.bas"
                $translationFile = Join-Path -Path $WorkDirectory -ChildPath "out.xml"
                $outputFile = Join-Path -Path $WorkDirectory -ChildPath "out.bas"
                $encoding = (New-Object -TypeName System.Text.UTF8Encoding -ArgumentList @($false))

                @"
Public DummyField As String
Public Sub DummyMethod()
    Debug.Print "_(NormalPlaceHolder_1)"
    Debug.Print "_(NormalPlaceHolder_2)"
End Sub
"@ | WrappedOutFile -Path $inputFile -Encoding $encoding -NoNewline
                @"
<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<translations>
<translation id="NormalPlaceHolder_1">This is normal string 1.</translation>
<translation id="NormalPlaceHolder_2">This is normal string 2.</translation>
</translations>
"@ | WrappedOutFile -Path $translationFile -Encoding $encoding -NoNewline
                Resolve-VBATranslationPlaceHolder -SourcePath $inputfile -SourceEncoding $encoding -TranslationPath $translationFile -DestinationPath $outputFile -DestinationEncoding $encoding

                WrappedGetContent -Path $outputFile -Raw | Should -Be (ReplaceLFToCRLF -Content @"
Public DummyField As String
Public Sub DummyMethod()
    Debug.Print "This is normal string 1."
    Debug.Print "This is normal string 2."
End Sub
"@)
                Remove-Item $inputFile -Force
                Remove-Item $translationFile -Force
                Remove-Item $outputFile -Force
            }
            It 'More than 1 placeholders inline' {
                $inputFile = Join-Path -Path $WorkDirectory -ChildPath "in.bas"
                $translationFile = Join-Path -Path $WorkDirectory -ChildPath "out.xml"
                $outputFile = Join-Path -Path $WorkDirectory -ChildPath "out.bas"
                $encoding = (New-Object -TypeName System.Text.UTF8Encoding -ArgumentList @($false))

                @"
Public DummyField As String
Public Sub DummyMethod()
    Debug.Print "_(NormalPlaceHolder_1)" & "_(NormalPlaceHolder_2)"
    Debug.Print "_(NormalPlaceHolder_3)": Debug.Print "_(NormalPlaceHolder_4)"
End Sub
"@ | WrappedOutFile -Path $inputFile -Encoding $encoding -NoNewline
                @"
<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<translations>
<translation id="NormalPlaceHolder_1">This is normal string 1.</translation>
<translation id="NormalPlaceHolder_2">This is normal string 2.</translation>
<translation id="NormalPlaceHolder_3">This is normal string 3.</translation>
<translation id="NormalPlaceHolder_4">This is normal string 4.</translation>
</translations>
"@ | WrappedOutFile -Path $translationFile -Encoding $encoding -NoNewline
                Resolve-VBATranslationPlaceHolder -SourcePath $inputfile -SourceEncoding $encoding -TranslationPath $translationFile -DestinationPath $outputFile -DestinationEncoding $encoding

                WrappedGetContent -Path $outputFile -Raw | Should -Be (ReplaceLFToCRLF -Content @"
Public DummyField As String
Public Sub DummyMethod()
    Debug.Print "This is normal string 1." & "This is normal string 2."
    Debug.Print "This is normal string 3.": Debug.Print "This is normal string 4."
End Sub
"@)
                Remove-Item $inputFile -Force
                Remove-Item $translationFile -Force
                Remove-Item $outputFile -Force
            }
            It 'Multiline translation' {
                $inputFile = Join-Path -Path $WorkDirectory -ChildPath "in.bas"
                $translationFile1 = Join-Path -Path $WorkDirectory -ChildPath "out1.xml"
                $outputFile1 = Join-Path -Path $WorkDirectory -ChildPath "out1.bas"
                $translationFile2 = Join-Path -Path $WorkDirectory -ChildPath "out2.xml"
                $outputFile2 = Join-Path -Path $WorkDirectory -ChildPath "out2.bas"
                $encoding = (New-Object -TypeName System.Text.UTF8Encoding -ArgumentList @($false))

                @"
Public DummyField As String
Public Sub DummyMethod()
    Debug.Print "_(NormalPlaceHolder_1)"
End Sub
"@ | WrappedOutFile -Path $inputFile -Encoding $encoding -NoNewline
                @"
<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<translations>
<translation id="NormalPlaceHolder_1">This is
normal
string 1.</translation>
</translations>
"@ | WrappedOutFile -Path $translationFile1 -Encoding $encoding -NoNewline
                Resolve-VBATranslationPlaceHolder -SourcePath $inputfile -SourceEncoding $encoding -TranslationPath $translationFile1 -DestinationPath $outputFile1 -DestinationEncoding $encoding

                WrappedGetContent -Path $outputFile1 -Raw | Should -Be (ReplaceLFToCRLF -Content @"
Public DummyField As String
Public Sub DummyMethod()
    Debug.Print "This is" & VbCrLf & "normal" & VbCrLf & "string 1."
End Sub
"@)

                @"
<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<translations>
<translation id="NormalPlaceHolder_1">
This is
normal
string 1.
</translation>
</translations>
"@ | WrappedOutFile -Path $translationFile2 -Encoding $encoding -NoNewline
                Resolve-VBATranslationPlaceHolder -SourcePath $inputfile -SourceEncoding $encoding -TranslationPath $translationFile2 -DestinationPath $outputFile2 -DestinationEncoding $encoding

                WrappedGetContent -Path $outputFile2 -Raw | Should -Be (WrappedGetContent -Path $outputFile1 -Raw)

                Remove-Item $inputFile -Force
                Remove-Item $translationFile1 -Force
                Remove-Item $outputFile1 -Force
                Remove-Item $translationFile2 -Force
                Remove-Item $outputFile2 -Force
            }
            It 'Translation not defined' {
                $inputFile = Join-Path -Path $WorkDirectory -ChildPath "in.bas"
                $translationFile = Join-Path -Path $WorkDirectory -ChildPath "out.xml"
                $outputFile = Join-Path -Path $WorkDirectory -ChildPath "out.bas"
                $warningStreamLog = Join-Path -Path $WorkDirectory -ChildPath "warning.log"
                $encoding = (New-Object -TypeName System.Text.UTF8Encoding -ArgumentList @($false))

                @"
Public DummyField As String
Public Sub DummyMethod()
    Debug.Print "_(NormalPlaceHolder_1)"
End Sub
"@ | WrappedOutFile -Path $inputFile -Encoding $encoding -NoNewline
                @"
<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<translations>
<translation id="NormalPlaceHolder_2">This is normal string 2.</translation>
<translation id="NormalPlaceHolder_3">This is normal string 3.</translation>
</translations>
"@ | WrappedOutFile -Path $translationFile -Encoding $encoding -NoNewline
                Resolve-VBATranslationPlaceHolder -SourcePath $inputfile -SourceEncoding $encoding -TranslationPath $translationFile -DestinationPath $outputFile -DestinationEncoding $encoding 3>$warningStreamLog
                WrappedGetContent -Path $warningStreamLog -Raw | Should -Be  "Translation id=`"(NormalPlaceHolder_1)`" is not defined.$([System.Environment]::NewLine)"
                WrappedGetContent -Path $outputFile -Raw | Should -Be (ReplaceLFToCRLF -Content @"
Public DummyField As String
Public Sub DummyMethod()
    Debug.Print "_(NormalPlaceHolder_1)"
End Sub
"@)
                Remove-Item $inputFile -Force
                Remove-Item $translationFile -Force
                Remove-Item $outputFile -Force
            }
            It 'Class description placeholder' {
                $inputFile = Join-Path -Path $WorkDirectory -ChildPath "in.bas"
                $translationFile = Join-Path -Path $WorkDirectory -ChildPath "out.xml"
                $outputFile = Join-Path -Path $WorkDirectory -ChildPath "out.bas"
                $encoding = (New-Object -TypeName System.Text.UTF8Encoding -ArgumentList @($false))

                @"
Attribute VB_Description = "_(ClassDescriptionPlaceHolder)"
Public DummyField As String
Public Sub DummyMethod()
End Sub
"@ | WrappedOutFile -Path $inputFile -Encoding $encoding -NoNewline
                @"
<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<translations>
<translation id="ClassDescriptionPlaceHolder">This is class description.</translation>
</translations>
"@ | WrappedOutFile -Path $translationFile -Encoding $encoding -NoNewline

                Resolve-VBATranslationPlaceHolder -SourcePath $inputfile -SourceEncoding $encoding -TranslationPath $translationFile -DestinationPath $outputFile -DestinationEncoding $encoding

                WrappedGetContent -Path $outputFile -Raw | Should -Be (ReplaceLFToCRLF -Content @"
Attribute VB_Description = "This is class description."
Public DummyField As String
Public Sub DummyMethod()
End Sub
"@)

                Remove-Item $inputFile -Force
                Remove-Item $translationFile -Force
                Remove-Item $outputFile -Force
            }
            It 'Multiline class description translation' {
                $inputFile = Join-Path -Path $WorkDirectory -ChildPath "in.bas"
                $translationFile1 = Join-Path -Path $WorkDirectory -ChildPath "out1.xml"
                $outputFile1 = Join-Path -Path $WorkDirectory -ChildPath "out1.bas"
                $translationFile2 = Join-Path -Path $WorkDirectory -ChildPath "out2.xml"
                $outputFile2 = Join-Path -Path $WorkDirectory -ChildPath "out2.bas"
                $encoding = (New-Object -TypeName System.Text.UTF8Encoding -ArgumentList @($false))

                @"
Attribute VB_Description = "_(ClassDescriptionPlaceHolder)"
Public DummyField As String
Public Sub DummyMethod()
End Sub
"@ | WrappedOutFile -Path $inputFile -Encoding $encoding -NoNewline
                @"
<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<translations>
<translation id="ClassDescriptionPlaceHolder">This is
class
description.</translation>
</translations>
"@ | WrappedOutFile -Path $translationFile1 -Encoding $encoding -NoNewline
                Resolve-VBATranslationPlaceHolder -SourcePath $inputfile -SourceEncoding $encoding -TranslationPath $translationFile1 -DestinationPath $outputFile1 -DestinationEncoding $encoding

                WrappedGetContent -Path $outputFile1 -Raw | Should -Be (ReplaceLFToCRLF -Content @"
Attribute VB_Description = "This is\n    class\n    description."
Public DummyField As String
Public Sub DummyMethod()
End Sub
"@)

                @"
<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<translations>
<translation id="ClassDescriptionPlaceHolder">
This is
class
description.
</translation>
</translations>
"@ | WrappedOutFile -Path $translationFile2 -Encoding $encoding -NoNewline
                Resolve-VBATranslationPlaceHolder -SourcePath $inputfile -SourceEncoding $encoding -TranslationPath $translationFile2 -DestinationPath $outputFile2 -DestinationEncoding $encoding

                WrappedGetContent -Path $outputFile2 -Raw | Should -Be (WrappedGetContent -Path $outputFile1 -Raw)

                Remove-Item $inputFile -Force
                Remove-Item $translationFile1 -Force
                Remove-Item $outputFile1 -Force
                Remove-Item $translationFile2 -Force
                Remove-Item $outputFile2 -Force
            }
            It 'Field description placeholder' {
                $inputFile = Join-Path -Path $WorkDirectory -ChildPath "in.bas"
                $translationFile = Join-Path -Path $WorkDirectory -ChildPath "out.xml"
                $outputFile = Join-Path -Path $WorkDirectory -ChildPath "out.bas"
                $encoding = (New-Object -TypeName System.Text.UTF8Encoding -ArgumentList @($false))

                @"
Public DummyField As String
Attribute DummyField.VB_VarDescription = "_(FieldDescriptionPlaceHolder)"
Public Sub DummyMethod()
End Sub
"@ | WrappedOutFile -Path $inputFile -Encoding $encoding -NoNewline
                @"
<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<translations>
<translation id="FieldDescriptionPlaceHolder">This is field description.</translation>
</translations>
"@ | WrappedOutFile -Path $translationFile -Encoding $encoding -NoNewline

                Resolve-VBATranslationPlaceHolder -SourcePath $inputfile -SourceEncoding $encoding -TranslationPath $translationFile -DestinationPath $outputFile -DestinationEncoding $encoding

                WrappedGetContent -Path $outputFile -Raw | Should -Be (ReplaceLFToCRLF -Content @"
Public DummyField As String
Attribute DummyField.VB_VarDescription = "This is field description."
Public Sub DummyMethod()
End Sub
"@)

                Remove-Item $inputFile -Force
                Remove-Item $translationFile -Force
                Remove-Item $outputFile -Force
            }
            It 'Multiline field description translation' {
                $inputFile = Join-Path -Path $WorkDirectory -ChildPath "in.bas"
                $translationFile1 = Join-Path -Path $WorkDirectory -ChildPath "out1.xml"
                $outputFile1 = Join-Path -Path $WorkDirectory -ChildPath "out1.bas"
                $translationFile2 = Join-Path -Path $WorkDirectory -ChildPath "out2.xml"
                $outputFile2 = Join-Path -Path $WorkDirectory -ChildPath "out2.bas"
                $encoding = (New-Object -TypeName System.Text.UTF8Encoding -ArgumentList @($false))

                @"
Public DummyField As String
Attribute DummyField.VB_VarDescription = "_(FieldDescriptionPlaceHolder)"
Public Sub DummyMethod()
End Sub
"@ | WrappedOutFile -Path $inputFile -Encoding $encoding -NoNewline
                @"
<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<translations>
<translation id="FieldDescriptionPlaceHolder">This is
field
description.</translation>
</translations>
"@ | WrappedOutFile -Path $translationFile1 -Encoding $encoding -NoNewline
                Resolve-VBATranslationPlaceHolder -SourcePath $inputfile -SourceEncoding $encoding -TranslationPath $translationFile1 -DestinationPath $outputFile1 -DestinationEncoding $encoding

                WrappedGetContent -Path $outputFile1 -Raw | Should -Be (ReplaceLFToCRLF -Content @"
Public DummyField As String
Attribute DummyField.VB_VarDescription = "This is\n    field\n    description."
Public Sub DummyMethod()
End Sub
"@)

                @"
<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<translations>
<translation id="FieldDescriptionPlaceHolder">
This is
field
description.
</translation>
</translations>
"@ | WrappedOutFile -Path $translationFile2 -Encoding $encoding -NoNewline
                Resolve-VBATranslationPlaceHolder -SourcePath $inputfile -SourceEncoding $encoding -TranslationPath $translationFile2 -DestinationPath $outputFile2 -DestinationEncoding $encoding

                WrappedGetContent -Path $outputFile2 -Raw | Should -Be (WrappedGetContent -Path $outputFile1 -Raw)

                Remove-Item $inputFile -Force
                Remove-Item $translationFile1 -Force
                Remove-Item $outputFile1 -Force
                Remove-Item $translationFile2 -Force
                Remove-Item $outputFile2 -Force
            }
            It 'Method description placeholder' {
                $inputFile = Join-Path -Path $WorkDirectory -ChildPath "in.bas"
                $translationFile = Join-Path -Path $WorkDirectory -ChildPath "out.xml"
                $outputFile = Join-Path -Path $WorkDirectory -ChildPath "out.bas"
                $encoding = (New-Object -TypeName System.Text.UTF8Encoding -ArgumentList @($false))

                @"
Public DummyField As String
Public Sub DummyMethod()
Attribute DummyMethod.VB_Description = "_(MethodDescriptionPlaceHolder)"
End Sub
"@ | WrappedOutFile -Path $inputFile -Encoding $encoding -NoNewline
                @"
<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<translations>
<translation id="MethodDescriptionPlaceHolder">This is method description.</translation>
</translations>
"@ | WrappedOutFile -Path $translationFile -Encoding $encoding -NoNewline

                Resolve-VBATranslationPlaceHolder -SourcePath $inputfile -SourceEncoding $encoding -TranslationPath $translationFile -DestinationPath $outputFile -DestinationEncoding $encoding

                WrappedGetContent -Path $outputFile -Raw | Should -Be (ReplaceLFToCRLF -Content @"
Public DummyField As String
Public Sub DummyMethod()
Attribute DummyMethod.VB_Description = "This is method description."
End Sub
"@)

                Remove-Item $inputFile -Force
                Remove-Item $translationFile -Force
                Remove-Item $outputFile -Force
            }
            It 'Multiline method description translation' {
                $inputFile = Join-Path -Path $WorkDirectory -ChildPath "in.bas"
                $translationFile1 = Join-Path -Path $WorkDirectory -ChildPath "out1.xml"
                $outputFile1 = Join-Path -Path $WorkDirectory -ChildPath "out1.bas"
                $translationFile2 = Join-Path -Path $WorkDirectory -ChildPath "out2.xml"
                $outputFile2 = Join-Path -Path $WorkDirectory -ChildPath "out2.bas"
                $encoding = (New-Object -TypeName System.Text.UTF8Encoding -ArgumentList @($false))

                @"
Public DummyField As String
Public Sub DummyMethod()
Attribute DummyMethod.VB_Description = "_(MethodDescriptionPlaceHolder)"
End Sub
"@ | WrappedOutFile -Path $inputFile -Encoding $encoding -NoNewline
                @"
<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<translations>
<translation id="MethodDescriptionPlaceHolder">This is
method
description.</translation>
</translations>
"@ | WrappedOutFile -Path $translationFile1 -Encoding $encoding -NoNewline
                Resolve-VBATranslationPlaceHolder -SourcePath $inputfile -SourceEncoding $encoding -TranslationPath $translationFile1 -DestinationPath $outputFile1 -DestinationEncoding $encoding

                WrappedGetContent -Path $outputFile1 -Raw | Should -Be (ReplaceLFToCRLF -Content @"
Public DummyField As String
Public Sub DummyMethod()
Attribute DummyMethod.VB_Description = "This is\n    method\n    description."
End Sub
"@)

                @"
<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<translations>
<translation id="MethodDescriptionPlaceHolder">
This is
method
description.
</translation>
</translations>
"@ | WrappedOutFile -Path $translationFile2 -Encoding $encoding -NoNewline
                Resolve-VBATranslationPlaceHolder -SourcePath $inputfile -SourceEncoding $encoding -TranslationPath $translationFile2 -DestinationPath $outputFile2 -DestinationEncoding $encoding

                WrappedGetContent -Path $outputFile2 -Raw | Should -Be (WrappedGetContent -Path $outputFile1 -Raw)

                Remove-Item $inputFile -Force
                Remove-Item $translationFile1 -Force
                Remove-Item $outputFile1 -Force
                Remove-Item $translationFile2 -Force
                Remove-Item $outputFile2 -Force
            }
            It 'All pattern' {
                $inputFile = Join-Path -Path $WorkDirectory -ChildPath "in.bas"
                $translationFile = Join-Path -Path $WorkDirectory -ChildPath "out.xml"
                $outputFile = Join-Path -Path $WorkDirectory -ChildPath "out.bas"
                $encoding = (New-Object -TypeName System.Text.UTF8Encoding -ArgumentList @($false))

                @"
Attribute VB_Description = "_(ClassDescriptionPlaceHolder)"
Public DummyField As String
Attribute DummyField.VB_VarDescription = "_(FieldDescriptionPlaceHolder)"
Public Sub DummyMethod()
Attribute DummyMethod.VB_Description = "_(MethodDescriptionPlaceHolder)"
    Debug.Print "_(NormalPlaceHolder_1)"
    Debug.Print "_(NormalPlaceHolder_2)" & "_(NormalPlaceHolder_3)"
    Debug.Print "_(NormalPlaceHolder_4)": Debug.Print "_(NormalPlaceHolder_5)"
    Debug.Print """_(NotPlaceHolder)"""
End Sub
"@ | WrappedOutFile -Path $inputFile -Encoding $encoding -NoNewline
                @"
<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<translations>
<translation id="ClassDescriptionPlaceHolder">This is class description.</translation>
<translation id="FieldDescriptionPlaceHolder">This is field description.</translation>
<translation id="MethodDescriptionPlaceHolder">This is method description.</translation>
<translation id="NormalPlaceHolder_1">This is normal string 1.</translation>
<translation id="NormalPlaceHolder_2">This is normal string 2.</translation>
<translation id="NormalPlaceHolder_3">This is normal string 3.</translation>
<translation id="NormalPlaceHolder_4">This is normal string 4.</translation>
<translation id="NormalPlaceHolder_5">This is normal string 5.</translation>
</translations>
"@ | WrappedOutFile -Path $translationFile -Encoding $encoding -NoNewline

                Resolve-VBATranslationPlaceHolder -SourcePath $inputfile -SourceEncoding $encoding -TranslationPath $translationFile -DestinationPath $outputFile -DestinationEncoding $encoding

                WrappedGetContent -Path $outputFile -Raw | Should -Be (ReplaceLFToCRLF -Content @"
Attribute VB_Description = "This is class description."
Public DummyField As String
Attribute DummyField.VB_VarDescription = "This is field description."
Public Sub DummyMethod()
Attribute DummyMethod.VB_Description = "This is method description."
    Debug.Print "This is normal string 1."
    Debug.Print "This is normal string 2." & "This is normal string 3."
    Debug.Print "This is normal string 4.": Debug.Print "This is normal string 5."
    Debug.Print """_(NotPlaceHolder)"""
End Sub
"@)

                Remove-Item $inputFile -Force
                Remove-Item $translationFile -Force
                Remove-Item $outputFile -Force
            }
            It 'File already exists with force' {
                $inputFile = Join-Path -Path $WorkDirectory -ChildPath "in.bas"
                $translationFile = Join-Path -Path $WorkDirectory -ChildPath "out.xml"
                $outputFile = Join-Path -Path $WorkDirectory -ChildPath "out.bas"
                $encoding = (New-Object -TypeName System.Text.UTF8Encoding -ArgumentList @($false))

                @"
"@ | WrappedOutFile -Path $inputFile -Encoding $encoding -NoNewline
                @"
<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<translations>
</translations>
"@ | WrappedOutFile -Path $translationFile -Encoding $encoding -NoNewline
                @"
"@ | WrappedOutFile -Path $outputFile -Encoding $encoding -NoNewline

                Resolve-VBATranslationPlaceHolder -SourcePath $inputfile -SourceEncoding $encoding -TranslationPath $translationFile -DestinationPath $outputFile -DestinationEncoding $encoding -Force

                Remove-Item $inputFile -Force
                Remove-Item $translationFile -Force
                Remove-Item $outputFile -Force
            }
            It 'Create a parent directory if it does not exist.' {
                $inputFile = Join-Path -Path $WorkDirectory -ChildPath "in.bas"
                $translationFile = Join-Path -Path $WorkDirectory -ChildPath "out.xml"
                $outputFile = Join-Path -Path $WorkDirectory -ChildPath "out\out.bas"
                $outputParent = Split-Path -Path $outputFile -Parent
                $encoding = (New-Object -TypeName System.Text.UTF8Encoding -ArgumentList @($false))

                @"
"@ | WrappedOutFile -Path $inputFile -Encoding $encoding -NoNewline
                @"
<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<translations>
</translations>
"@ | WrappedOutFile -Path $translationFile -Encoding $encoding -NoNewline
                Resolve-VBATranslationPlaceHolder -SourcePath $inputfile -SourceEncoding $encoding -TranslationPath $translationFile -DestinationPath $outputFile -DestinationEncoding $encoding
                Test-Path -Path $outputParent | Should -BeTrue

                Remove-Item $inputFile -Force
                Remove-Item $translationFile -Force
                Remove-Item $outputFile -Force
                Remove-Item $outputParent -Force
            }
        }
        Context 'Exception' {
            It 'Invalid translation format' {
                $inputFile = Join-Path -Path $WorkDirectory -ChildPath "in.bas"
                $translationFile = Join-Path -Path $WorkDirectory -ChildPath "out.xml"
                $outputFile = Join-Path -Path $WorkDirectory -ChildPath "out.bas"
                $encoding = (New-Object -TypeName System.Text.UTF8Encoding -ArgumentList @($false))

                @"
"@ | WrappedOutFile -Path $inputFile -Encoding $encoding -NoNewline
                @"
<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<translationss>
<translation id="NormalPlaceHolder_1">This is normal string 1.</translation>
</translationss>
"@ | WrappedOutFile -Path $translationFile -Encoding $encoding -NoNewline

                {Resolve-VBATranslationPlaceHolder -SourcePath $inputfile -SourceEncoding $encoding -TranslationPath $translationFile -DestinationPath $outputFile -DestinationEncoding $encoding} | Should -Throw

                @"
"@ | WrappedOutFile -Path $inputFile -Encoding $encoding -NoNewline
@"
<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<translations>
<translationn id="NormalPlaceHolder_1">This is normal string 1.</translationn>
</translations>
"@ | WrappedOutFile -Path $translationFile -Encoding $encoding -NoNewline

                {Resolve-VBATranslationPlaceHolder -SourcePath $inputfile -SourceEncoding $encoding -TranslationPath $translationFile -DestinationPath $outputFile -DestinationEncoding $encoding} | Should -Throw

                @"
<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<translations>
<translation idd="NormalPlaceHolder_1">This is normal string 1.</translation>
</translations>
"@ | WrappedOutFile -Path $translationFile -Encoding $encoding -NoNewline

                {Resolve-VBATranslationPlaceHolder -SourcePath $inputfile -SourceEncoding $encoding -TranslationPath $translationFile -DestinationPath $outputFile -DestinationEncoding $encoding} | Should -Throw

                Remove-Item $inputFile -Force
                Remove-Item $translationFile -Force
                if(Test-Path -Path $outputFile)
                {
                    Remove-Item $outputFile -Force
                }
            }
            It 'File already exists without force' {
                $inputFile = Join-Path -Path $WorkDirectory -ChildPath "in.bas"
                $translationFile = Join-Path -Path $WorkDirectory -ChildPath "out.xml"
                $outputFile = Join-Path -Path $WorkDirectory -ChildPath "out.bas"
                $encoding = (New-Object -TypeName System.Text.UTF8Encoding -ArgumentList @($false))

                @"
"@ | WrappedOutFile -Path $inputFile -Encoding $encoding -NoNewline
                @"
<?xml version="1.0" encoding="utf-8" standalone="yes"?>
</translations>
"@ | WrappedOutFile -Path $translationFile -Encoding $encoding -NoNewline
                @"
"@ | WrappedOutFile -Path $outputFile -Encoding $encoding -NoNewline

                {Resolve-VBATranslationPlaceHolder -SourcePath $inputfile -SourceEncoding $encoding -TranslationPath $translationFile -DestinationPath $outputFile -DestinationEncoding $encoding} | Should -Throw

                Remove-Item $inputFile -Force
                Remove-Item $translationFile -Force
                Remove-Item $outputFile -Force
            }
        }
    }
}
