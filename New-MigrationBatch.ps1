Param(
    [String]$BatchName = 'Test Migration',

    [String]$MigrationEndpoint = 'prudhommewtf_abcloudeuh_3257',

    [String]$CsvFilePath = "$PSScriptRoot\users.csv",

    [String]$TargetDeliveryDomain = 'prudhommewtf.onmicrosoft.com'
)

New-MigrationBatch `
    -Name $BatchName `
    -SourceEndpoint $MigrationEndpoint `
    -CSVData ([System.IO.File]::ReadAllBytes($CsvFilePath)) `
    -TargetDeliveryDomain $TargetDeliveryDomain `
    -Autostart `
    -AutoComplete