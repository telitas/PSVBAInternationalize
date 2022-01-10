# Copyright (c) 2022 telitas
# This file is released under the MIT License.
# See the LICENSE file or https://opensource.org/licenses/mit-license.php for details.

function Export-VBATranslationPlaceHolder
{
    <#
    .EXTERNALHELP ..\PSVBAInternationalize-help.xml
    #>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)][string]$SourcePath,
        [System.Text.Encoding]$SourceEncoding = (New-Object -TypeName System.Text.UTF8Encoding -ArgumentList @($false)),
        [Parameter(Mandatory=$true)][string]$DestinationPath,
        [switch]$Force
    )
    if(-not $Force -and (Test-Path -Path $DestinationPath))
    {
        Write-Error "Destination file `"$DestinationPath` is already exists." -ErrorAction Stop
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
    Set-Variable -Name PlaceHolderPattern -Value (New-Object -TypeName regex -ArgumentList @('(?<!"+)"_\((?<id>[A-Za-z0-9_\-]+?)\)"(?!"+)', [System.Text.RegularExpressions.RegexOptions]::Compiled)) -Option ReadOnly

    [string[]]$placeHolders = @()
    WrappedGetContent -Path $SourcePath -Encoding $SourceEncoding | ForEach-Object -Process {
        $placeHolders += $PlaceHolderPattern.Matches($_) | ForEach-Object -Process {$_.Groups["id"].Value}
    }
    $document = New-Object -TypeName System.Xml.Linq.XDocument -ArgumentList @(
        (New-Object -TypeName System.Xml.Linq.XDeclaration -ArgumentList @("1.0", "utf-8", "yes")),
        (New-Object -TypeName System.Xml.Linq.XElement -ArgumentList @(
            "translations",
            ([System.Xml.Linq.XElement[]]($placeHolders | Sort-Object -Unique | ForEach-Object -Process {New-Object -TypeName System.Xml.Linq.XElement -ArgumentList @("translation", (New-Object -TypeName System.Xml.Linq.XAttribute -ArgumentList @("id", $_)), "")}))
        ))
    )

    [System.Xml.XmlWriterSettings]$setting = New-Object -TypeName System.Xml.XmlWriterSettings
    $setting.Encoding =  New-Object -TypeName System.Text.UTF8Encoding -ArgumentList @($false)
    $setting.Indent = $true
    $setting.IndentChars = ""
    $setting.NewLineChars = "`n"
    $writer = [System.Xml.XmlWriter]::Create($DestinationPath, $setting)
    try
    {
        $document.Save($writer)
    }
    finally
    {
        $writer.Close()
    }
}
