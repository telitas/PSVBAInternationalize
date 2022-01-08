Get-ChildItem -Path "$($PSScriptRoot)\src\*.ps1" | ForEach-Object -Process {. $_.FullName}
