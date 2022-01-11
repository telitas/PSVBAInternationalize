# Copyright (c) 2022 telitas
# This file is released under the MIT License.
# See the LICENSE file or https://opensource.org/licenses/mit-license.php for details.

function Resolve-VBATranslationPlaceHolder
{
    <#
    .EXTERNALHELP ..\PSVBAInternationalize-help.xml
    #>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)][string]$SourcePath,
        [System.Text.Encoding]$SourceEncoding = (New-Object -TypeName System.Text.UTF8Encoding -ArgumentList @($false)),
        [Parameter(Mandatory=$true)][string]$TranslationPath,
        [Parameter(Mandatory=$true)][string]$DestinationPath,
        [System.Text.Encoding]$DestinationEncoding = (New-Object -TypeName System.Text.UTF8Encoding -ArgumentList @($false)),
        [switch]$Force
    )
    if(-not $Force -and (Test-Path -Path $DestinationPath))
    {
        Write-Error "Destination file `"$DestinationPath` is already exists." -ErrorAction Stop
    }
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
    $UIMessages = data {}
    try
    {
        Import-LocalizedData -BindingVariable UIMessages -ErrorAction Stop
    }
    catch
    {
        Import-LocalizedData -BindingVariable UIMessages -UICulture 'en-US'
    }
    Set-Variable -Name SupportPowerShellVersion -Value 3 -Option ReadOnly
    if($PSVersionTable.PSVersion.Major -lt $SupportPowerShellVersion)
    {
        Write-Warning -Message $UIMessages.ThisVersionIsNotSupported.Replace('$SupportPowerShellVersion', $SupportPowerShellVersion)
    }
    if($PSVersionTable.PSVersion.Major -le 5)
    {
        [System.Reflection.Assembly]::LoadWithPartialName("System.Xml.Linq") | Out-Null
    }

    Set-Variable -Name Translation -Value @{} -Option ReadOnly

    [System.Xml.Schema.XmlSchemaSet]$schemas = New-Object -TypeName System.Xml.Schema.XmlSchemaSet
    $schemas.Add("", (Join-Path -Path $PSScriptRoot -ChildPath "translations.xsd")) > $null
    [System.Xml.Linq.XDocument]$document = [System.Xml.Linq.XDocument]::Load($TranslationPath)
    [bool]$script:validated = $true
    [System.Xml.Schema.Extensions]::Validate($document, $schemas, {$script:validated = $false})
    if(-not $script:validated)
    {
        Write-Error -Message "$($TranslationPath) is invalid format." -ErrorAction Stop
    }

    Set-Variable -Name PlaceHolderPattern -Value (New-Object -TypeName regex -ArgumentList @('(?<!"+)"_\((?<id>[A-Za-z0-9_\-]+?)\)"(?!"+)', [System.Text.RegularExpressions.RegexOptions]::Compiled)) -Option ReadOnly
    Set-Variable -Name DescriptionPlaceHolderPattern -Value (New-Object -TypeName regex -ArgumentList @("^ *Attribute +(?:[A-Za-z0-9_]+\.)?VB_(?:Var)?Description *= *$($PlaceHolderPattern.ToString()) *$", [System.Text.RegularExpressions.RegexOptions]::Compiled)) -Option ReadOnly

    ([xml](WrappedGetContent -Path $TranslationPath)).translations.translation | ForEach-Object -Process {
        if($null -ne $_){
            $trans = $_."#text"
            if($null -eq $trans)
            {
                $trans = ""
            }
            if($trans.StartsWith("`r`n"))
            {
                $trans = $trans.SubString(2)
            }
            elseif($trans.StartsWith("`r") -or $trans.StartsWith("`n"))
            {
                $trans = $trans.SubString(1)
            }
            if($trans.EndsWith("`r`n"))
            {
                $trans = $trans.SubString(0, $trans.Length - 2)
            }
            elseif($trans.EndsWith("`r") -or $trans.EndsWith("`n"))
            {
                $trans = $trans.SubString(0, $trans.Length - 1)
            }
            $Translation[$_.id] = $trans
        }
    }
    $destinationParent = Split-Path -Path $DestinationPath -Parent
    if(-not (Test-Path -Path $destinationParent))
    {
        New-Item -ItemType Directory -Path $destinationParent | Out-Null
    }
    (
        WrappedGetContent -Path $SourcePath -Encoding $SourceEncoding | ForEach-Object -Process {
            $content = $_
            if($content -cmatch $DescriptionPlaceHolderPattern)
            {
                $id=$Matches['id']
                if($Translation.Keys -notcontains $id)
                {
                    Write-Warning -Message "Translation id=`"($id)`" is not defined."
                    $content
                }
                else
                {
                    $content.Replace("_($($id))", ($Translation[$id] -replace "`n", '\n    '))
                }
            }
            else
            {
                $placeHolders = $PlaceHolderPattern.Matches($content)
                if($placeHolders.Count -gt 0)
                {
                    $placeHolders | ForEach-Object -Process {
                        $id=$_.Groups["id"].Value
                        if($Translation.Keys -notcontains $id)
                        {
                            Write-Warning -Message "Translation id=`"($id)`" is not defined."
                        }
                        else
                        {
                            $content = $content.Replace("_($($id))", ($Translation[$id] -replace "`n", '" & VbCrLf & "'))
                        }
                    }
                    $content
                }
                else
                {
                    $content
                }
            }
        }
    ) -join "`r`n" | WrappedOutFile -Path $DestinationPath -Encoding $DestinationEncoding -NoNewline
}
