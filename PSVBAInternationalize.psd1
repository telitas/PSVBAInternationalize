@{
    RootModule = 'PSVBAInternationalize.psm1'
    ModuleVersion = '$Version'
    GUID = 'ecb8e4c4-3b89-4cc9-a71e-e64303d35b5d'
    Author = 'telitas'
    CompanyName = 'Unknown'
    Copyright = '(c) 2022 telitas'
    Description = 'VBA source files internationalize tool.'
    PowerShellVersion = '3.0'
    FunctionsToExport = @(
        'Export-VBATranslationPlaceHolder',
        'Resolve-VBATranslationPlaceHolder'
    )
    CmdletsToExport = @()
    VariablesToExport = '*'
    AliasesToExport = @()
    PrivateData = @{
        PSData = @{
            Tags = @('VBA', 'Windows', 'MacOS', 'Linux')
            License = 'MIT'
            ProjectUri = 'https://github.com/telitas/PSVBAInternationalize'
        }
    }
}

