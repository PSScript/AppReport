@{
    RootModule           = 'AppReport.psm1'
    ModuleVersion        = '1.0.0'
    GUID                 = '6f4b2c1e-8a3d-4e7b-9c5f-2d1a0b7e3c94'
    Author               = 'Jan Hübener'
    CompanyName          = 'PSScript'
    Description          = 'Inventory of mail / M365 workload app registrations: permissions, scope (ApplicationAccessPolicy, Exchange RBAC for Applications), credential expiry and EWS org / mailbox exceptions. CSV and XLSX export.'
    PowerShellVersion    = '5.1'
    RequiredModules      = @('Microsoft.Graph.Authentication')
    FunctionsToExport    = @(
        'Set-AppReportConfig', 'Get-AppReportConfig', 'Connect-AppReport',
        'Get-AppReportInventory', 'Get-AppReportAccessPolicy', 'Get-AppReportExchangeRbac',
        'Get-AppReportEwsConfig', 'Get-AppReportScope', 'Invoke-AppReport'
    )
    CmdletsToExport      = @()
    VariablesToExport    = @()
    AliasesToExport      = @()
    PrivateData          = @{
        PSData = @{
            Tags         = @('Exchange', 'ExchangeOnline', 'EntraID', 'Graph', 'AppRegistration', 'EWS', 'RBAC', 'ApplicationAccessPolicy', 'Report')
            ProjectUri   = 'https://github.com/PSScript/AppReport'
            ExternalModuleDependencies = @('ExchangeOnlineManagement', 'ImportExcel')
        }
    }
}
