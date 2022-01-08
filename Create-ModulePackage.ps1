if($PSVersionTable.PSVersion.Major -lt 7)
{
    Write-Error -Message "This script supports PowerShell 7 or later." -ErrorAction Stop
}
Set-Variable -Name PackageName -Value "PSVBAInternationalize" -Option ReadOnly
Set-Variable -Name PackagePath -Value (Join-Path -Path $PSScriptRoot -ChildPath $PackageName) -Option ReadOnly
Set-Variable -Name FilesMapping -Value @{
    "$($PackageName).psm1"=$null
    "$($PackageName).psd1"=$null
    "src"=$null
    "LICENSE.txt"=$null
} -Option ReadOnly
if((git tag --list --contains HEAD) -match "^v(?<version>[0-9]+\.[0-9]+\.[0-9])")
{
    $version = $Matches["version"]
}
else
{
    Write-Warning "The current Commit does not contain versioning tag."
    $version = "0.0.1"
}
if(Test-Path -Path $PackagePath){
    Write-Warning -Message "Old package was removed."
    Remove-Item -Path $PackagePath -Recurse
}
New-Item -ItemType Directory -Path $PackagePath > $null

$FilesMapping.Keys | ForEach-Object -Process {
    switch($_)
    {
        "$($PackageName).psd1"
        {
            (Get-Content -Path (Join-Path -Path $PSScriptRoot -ChildPath $_) -Raw) -replace 'ModuleVersion = ''\$Version''',  "ModuleVersion = '$($version)'" | Out-File -FilePath (Join-Path -Path $PackagePath -ChildPath $(if($null -eq $FilesMapping[$_]){$_}else{$FilesMapping[$_]})) -NoNewline
        }
        default
        {
            Copy-Item -Path (Join-Path -Path $PSScriptRoot -ChildPath $_) -Destination (Join-Path -Path $PackagePath -ChildPath $(if($null -eq $FilesMapping[$_]){$_}else{$FilesMapping[$_]})) -Recurse
        }
    }
}

Set-Variable -Name DocumetsPath -Value (Join-Path -Path $PSScriptRoot -ChildPath "docs") -Option ReadOnly
New-ExternalHelp -Path $DocumetsPath -OutputPath $PackagePath > $null
Get-ChildItem -Path $DocumetsPath -Attributes Directory -Recurse | ForEach-Object -Process {
    New-ExternalHelp -Path $_.FullName -OutputPath $_.FullName.Replace($DocumetsPath, $PackagePath) > $null
}
Test-ModuleManifest -Path (Join-Path -Path $PackagePath -ChildPath "$($PackageName).psd1")

Write-Output "New package was created."