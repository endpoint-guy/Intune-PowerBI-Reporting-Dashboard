$ErrorActionPreference = "Stop"

$DataPath = "./Data"
$SourceUrl = "https://api.datafornerds.io/v2/microsoft/windows-update-history.json"

if (-not (Test-Path -Path $DataPath)) {
    New-Item -Path $DataPath -ItemType Directory -Force | Out-Null
}

function Save-JsonArray {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$InputObject,

        [Parameter(Mandatory = $true)]
        [string]$Path,

        [int]$Depth = 10
    )

    if ($null -eq $InputObject -or $InputObject.Count -eq 0) {
        "[]" | Set-Content -Path $Path -Encoding UTF8
    }
    else {
        $InputObject |
            ConvertTo-Json -Depth $Depth -AsArray |
            Set-Content -Path $Path -Encoding UTF8
    }
}

function Save-JsonObject {
    param(
        [Parameter(Mandatory = $true)]
        [object]$InputObject,

        [Parameter(Mandatory = $true)]
        [string]$Path,

        [int]$Depth = 10
    )

    $InputObject |
        ConvertTo-Json -Depth $Depth |
        Set-Content -Path $Path -Encoding UTF8
}

function Get-ProductName {
    param(
        [object]$MajorVersion
    )

    switch ([string]$MajorVersion) {
        "11" { return "Windows 11" }
        "10" { return "Windows 10" }
        default { return "Windows Server" }
    }
}

function Convert-ToBooleanSafe {
    param(
        [object]$Value
    )

    if ($null -eq $Value) {
        return $false
    }

    try {
        return [System.Convert]::ToBoolean($Value)
    }
    catch {
        return $false
    }
}

function Convert-ToDateTimeSafe {
    param(
        [object]$Value
    )

    if ($null -eq $Value) {
        return $null
    }

    $StringValue = [string]$Value

    if (:IsNullOrWhiteSpace($StringValue)) {
        return $null
    }

    try {
        return [datetime]$StringValue
    }
    catch {
        return $null
    }
}

Write-Host "Downloading Windows Update history from public API..."
Write-Host "Source: $SourceUrl"

$Response = Invoke-RestMethod -Uri $SourceUrl -Method Get

if ($null -eq $Response) {
    throw "No response was returned from the API."
}

if ($null -eq $Response.data) {
    throw "No data property was found in the API response."
}

$ApiRecords = @($Response.data)

if ($ApiRecords.Count -eq 0) {
    throw "The API returned zero update records."
}

Write-Host "Records received from API: $($ApiRecords.Count)"

$GeneratedUtc = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")

$AllUpdates = @(
    foreach ($Update in $ApiRecords) {
        $OSBuildValue = [string]$Update.OSBuild
        $BuildBranch = $null
        $UBR = $null

        if (-not :IsNullOrWhiteSpace($OSBuildValue)) {
            $BuildParts = $OSBuildValue -split "\."

            if ($BuildParts.Count -eq 2) {
                try {
                    $BuildBranch = [int]$BuildParts[0]
                    $UBR = [int]$BuildParts[1]
                }
                catch {
                    $BuildBranch = $null
                    $UBR = $null
                }
            }
        }

        $ReleaseDateValue = Convert-ToDateTimeSafe -Value $Update.ReleaseDate
        $IsExpiredValue = Convert-ToBooleanSafe -Value $Update.IsExpired
        $ProductName = Get-ProductName -MajorVersion $Update.MajorVersion

        [PSCustomObject]@{
            OSType          = [string]$Update.OSType
            Product         = $ProductName
            MajorVersion    = $Update.MajorVersion
            Version         = [string]$Update.WindowsVersion
            KB              = [string]$Update.KBNumber
            Build           = [string]$Update.OSBuild
            FullVersion     = [string]$Update.FullVersion
            BuildBranch     = $BuildBranch
            UBR             = $UBR
            ReleaseDate     = if ($null -ne $ReleaseDateValue) { $ReleaseDateValue.ToString("yyyy-MM-dd") } else { $null }
            ReleaseDateSort = $ReleaseDateValue
            ReleaseType     = [string]$Update.ReleaseType
            IsExpired       = $IsExpiredValue
            ArticleUrl      = [string]$Update.ArticleUrl
            Source          = "DataForNerds"
            SourceUrl       = $SourceUrl
            GeneratedUtc    = $GeneratedUtc
        }
    }
)

$AllUpdates = @(
    $AllUpdates |
        Where-Object {
            -not :IsNullOrWhiteSpace($_.KB) -and
            -not :IsNullOrWhiteSpace($_.Build) -and
            -not :IsNullOrWhiteSpace($_.FullVersion) -and
            $null -ne $_.BuildBranch -and
            $null -ne $_.UBR -and
            $null -ne $_.ReleaseDateSort
        } |
        Sort-Object OSType, Product, Version, BuildBranch, UBR, ReleaseDateSort -Unique
)

Write-Host "Valid update records after normalization: $($AllUpdates.Count)"

$ClientUpdates = @(
    $AllUpdates |
        Where-Object {
            $_.OSType -eq "Client"
        } |
        Sort-Object Product, Version, BuildBranch, UBR, ReleaseDateSort -Unique
)

Write-Host "Client update records: $($ClientUpdates.Count)"

$NewestClientUpdate = $ClientUpdates |
    Sort-Object ReleaseDateSort, UBR -Descending |
    Select-Object -First 1

if ($null -ne $NewestClientUpdate) {
    Write-Host "Newest client update found:"
    Write-Host "Product: $($NewestClientUpdate.Product)"
    Write-Host "Version: $($NewestClientUpdate.Version)"
    Write-Host "KB: $($NewestClientUpdate.KB)"
    Write-Host "Build: $($NewestClientUpdate.Build)"
    Write-Host "FullVersion: $($NewestClientUpdate.FullVersion)"
    Write-Host "ReleaseDate: $($NewestClientUpdate.ReleaseDate)"
    Write-Host "ReleaseType: $($NewestClientUpdate.ReleaseType)"
    Write-Host "IsExpired: $($NewestClientUpdate.IsExpired)"
}

$ExportProperties = @(
    "OSType",
    "Product",
    "MajorVersion",
    "Version",
    "KB",
    "Build",
    "FullVersion",
    "BuildBranch",
    "UBR",
    "ReleaseDate",
    "ReleaseType",
    "IsExpired",
    "ArticleUrl",
    "Source",
    "SourceUrl",
    "GeneratedUtc"
)

$WindowsUpdateCatalog = @(
    $AllUpdates |
        Select-Object -Property $ExportProperties
)

Save-JsonArray `
    -InputObject $WindowsUpdateCatalog `
    -Path "$DataPath/WindowsUpdateCatalog.json" `
    -Depth 10

Write-Host "Saved WindowsUpdateCatalog.json"

$ClientWindowsUpdateCatalog = @(
    $ClientUpdates |
        Select-Object -Property $ExportProperties
)

Save-JsonArray `
    -InputObject $ClientWindowsUpdateCatalog `
    -Path "$DataPath/ClientWindowsUpdateCatalog.json" `
    -Depth 10

Write-Host "Saved ClientWindowsUpdateCatalog.json"

$LatestQualityUpdates = @(
    $ClientUpdates |
        Where-Object {
            $_.IsExpired -eq $false -and
            $_.ReleaseType -notmatch "Preview"
        } |
        Group-Object Product, Version, BuildBranch |
        ForEach-Object {
            $_.Group |
                Sort-Object ReleaseDateSort, UBR -Descending |
                Select-Object -First 1
        } |
        Sort-Object Product, Version, BuildBranch |
        Select-Object -Property $ExportProperties
)

Save-JsonArray `
    -InputObject $LatestQualityUpdates `
    -Path "$DataPath/LatestQualityUpdates.json" `
    -Depth 10

Write-Host "Saved LatestQualityUpdates.json"
Write-Host "Latest quality update records: $($LatestQualityUpdates.Count)"

$PreviewUpdates = @(
    $ClientUpdates |
        Where-Object {
            $_.ReleaseType -match "Preview"
        } |
        Sort-Object Product, Version, BuildBranch, ReleaseDateSort, UBR |
        Select-Object -Property $ExportProperties
)

Save-JsonArray `
    -InputObject $PreviewUpdates `
    -Path "$DataPath/PreviewUpdates.json" `
    -Depth 10

Write-Host "Saved PreviewUpdates.json"
Write-Host "Preview update records: $($PreviewUpdates.Count)"

$ReleaseTypeSummary = @(
    $ClientUpdates |
        Group-Object ReleaseType |
        ForEach-Object {
            [PSCustomObject]@{
                ReleaseType = $_.Name
                Count       = $_.Count
            }
        } |
        Sort-Object ReleaseType
)

Save-JsonArray `
    -InputObject $ReleaseTypeSummary `
    -Path "$DataPath/ReleaseTypeSummary.json" `
    -Depth 5

Write-Host "Saved ReleaseTypeSummary.json"

$Metadata = $Response.metadata

$SourceInfo = [PSCustomObject]@{
    Name                     = "DataForNerds Windows Update History"
    Url                      = $SourceUrl
    GeneratedUtc             = $GeneratedUtc
    TotalApiRecords          = $ApiRecords.Count
    ValidNormalizedRecords   = $AllUpdates.Count
    ClientRecordCount        = $ClientUpdates.Count
    LatestQualityRecordCount = $LatestQualityUpdates.Count
    PreviewRecordCount       = $PreviewUpdates.Count
    NewestReleaseDate        = if ($null -ne $NewestClientUpdate) { $NewestClientUpdate.ReleaseDate } else { $null }
    NewestKB                 = if ($null -ne $NewestClientUpdate) { $NewestClientUpdate.KB } else { $null }
    NewestFullVersion        = if ($null -ne $NewestClientUpdate) { $NewestClientUpdate.FullVersion } else { $null }
    ApiProvider              = if ($null -ne $Metadata) { $Metadata.provider } else { $null }
    ApiVersion               = if ($null -ne $Metadata) { $Metadata.apiVersion } else { $null }
    ApiDataset               = if ($null -ne $Metadata) { $Metadata.dataset } else { $null }
    ApiLastModified          = if ($null -ne $Metadata) { $Metadata.lastModified } else { $null }
    ApiLastCollected         = if ($null -ne $Metadata) { $Metadata.lastCollected } else { $null }
    Description              = "Public Windows update reference data used for Power BI reporting."
}

Save-JsonObject `
    -InputObject $SourceInfo `
    -Path "$DataPath/UpdateSources.json" `
    -Depth 10

Write-Host "Saved UpdateSources.json"

Write-Host "Windows update reference data generation completed successfully."
