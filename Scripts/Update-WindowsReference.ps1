$ErrorActionPreference = "Stop"

$DataPath = "./Data"
$SourceUrl = "https://api.datafornerds.io/v2/microsoft/windows-update-history.json"

if (-not (Test-Path -Path $DataPath)) {
    New-Item -Path $DataPath -ItemType Directory -Force | Out-Null
}

function Test-StringHasValue {
    param(
        [object]$Value
    )

    if ($null -eq $Value) {
        return $false
    }

    $StringValue = "$Value"

    if ($StringValue.Trim().Length -eq 0) {
        return $false
    }

    return $true
}

function Save-JsonArray {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$InputObject,

        [Parameter(Mandatory = $true)]
        [string]$Path,

        [int]$Depth = 10
    )

    if ($null -eq $InputObject -or $InputObject.Count -eq 0) {
        "[]" | Set-Content -Path $Path -Encoding UTF8
        return
    }

    $InputObject |
        ConvertTo-Json -Depth $Depth -AsArray |
        Set-Content -Path $Path -Encoding UTF8
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

    $MajorVersionText = "$MajorVersion"

    if ($MajorVersionText -eq "11") {
        return "Windows 11"
    }

    if ($MajorVersionText -eq "10") {
        return "Windows 10"
    }

    return "Windows Server"
}

function Convert-ToBooleanSafe {
    param(
        [object]$Value
    )

    if ($null -eq $Value) {
        return $false
    }

    if ($Value -is [bool]) {
        return $Value
    }

    $StringValue = "$Value".Trim().ToLowerInvariant()

    if ($StringValue -eq "true" -or $StringValue -eq "1" -or $StringValue -eq "yes") {
        return $true
    }

    return $false
}

function Convert-ToDateTimeSafe {
    param(
        [object]$Value
    )

    if (-not (Test-StringHasValue -Value $Value)) {
        return $null
    }

    $Result = $null

    try {
        $Result = Get-Date -Date "$Value"
    }
    catch {
        $Result = $null
    }

    return $Result
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
        $OSBuildValue = "$($Update.OSBuild)"
        $BuildBranch = $null
        $UBR = $null

        if (Test-StringHasValue -Value $OSBuildValue) {
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

        $ReleaseDateString = $null

        if ($null -ne $ReleaseDateValue) {
            $ReleaseDateString = $ReleaseDateValue.ToString("yyyy-MM-dd")
        }

        [PSCustomObject]@{
            OSType          = "$($Update.OSType)"
            Product         = $ProductName
            MajorVersion    = $Update.MajorVersion
            Version         = "$($Update.WindowsVersion)"
            KB              = "$($Update.KBNumber)"
            Build           = "$($Update.OSBuild)"
            FullVersion     = "$($Update.FullVersion)"
            BuildBranch     = $BuildBranch
            UBR             = $UBR
            ReleaseDate     = $ReleaseDateString
            ReleaseDateSort = $ReleaseDateValue
            ReleaseType     = "$($Update.ReleaseType)"
            IsExpired       = $IsExpiredValue
            ArticleUrl      = "$($Update.ArticleUrl)"
            Source          = "DataForNerds"
            SourceUrl       = $SourceUrl
            GeneratedUtc    = $GeneratedUtc
        }
    }
)

$AllUpdates = @(
    $AllUpdates |
        Where-Object {
            (Test-StringHasValue -Value $_.KB) -and
            (Test-StringHasValue -Value $_.Build) -and
            (Test-StringHasValue -Value $_.FullVersion) -and
            ($null -ne $_.BuildBranch) -and
            ($null -ne $_.UBR) -and
            ($null -ne $_.ReleaseDateSort)
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

Save-JsonArray -InputObject $WindowsUpdateCatalog -Path "$DataPath/WindowsUpdateCatalog.json" -Depth 10

Write-Host "Saved WindowsUpdateCatalog.json"

$ClientWindowsUpdateCatalog = @(
    $ClientUpdates |
        Select-Object -Property $ExportProperties
)

Save-JsonArray -InputObject $ClientWindowsUpdateCatalog -Path "$DataPath/ClientWindowsUpdateCatalog.json" -Depth 10

Write-Host "Saved ClientWindowsUpdateCatalog.json"

$LatestQualityUpdates = @(
    $ClientUpdates |
        Where-Object {
            ($_.IsExpired -eq $false) -and
            ($_.ReleaseType -notmatch "Preview")
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

Save-JsonArray -InputObject $LatestQualityUpdates -Path "$DataPath/LatestQualityUpdates.json" -Depth 10

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

Save-JsonArray -InputObject $PreviewUpdates -Path "$DataPath/PreviewUpdates.json" -Depth 10

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

Save-JsonArray -InputObject $ReleaseTypeSummary -Path "$DataPath/ReleaseTypeSummary.json" -Depth 5

Write-Host "Saved ReleaseTypeSummary.json"

$Metadata = $Response.metadata

$ApiProvider = $null
$ApiVersion = $null
$ApiDataset = $null
$ApiLastModified = $null
$ApiLastCollected = $null

if ($null -ne $Metadata) {
    $ApiProvider = "$($Metadata.provider)"
    $ApiVersion = "$($Metadata.apiVersion)"
    $ApiDataset = "$($Metadata.dataset)"
    $ApiLastModified = "$($Metadata.lastModified)"
    $ApiLastCollected = "$($Metadata.lastCollected)"
}

$NewestReleaseDate = $null
$NewestKB = $null
$NewestFullVersion = $null

if ($null -ne $NewestClientUpdate) {
    $NewestReleaseDate = $NewestClientUpdate.ReleaseDate
    $NewestKB = $NewestClientUpdate.KB
    $NewestFullVersion = $NewestClientUpdate.FullVersion
}

$SourceInfo = [PSCustomObject]@{
    Name                     = "DataForNerds Windows Update History"
    Url                      = $SourceUrl
    GeneratedUtc             = $GeneratedUtc
    TotalApiRecords          = $ApiRecords.Count
    ValidNormalizedRecords   = $AllUpdates.Count
    ClientRecordCount        = $ClientUpdates.Count
    LatestQualityRecordCount = $LatestQualityUpdates.Count
    PreviewRecordCount       = $PreviewUpdates.Count
    NewestReleaseDate        = $NewestReleaseDate
    NewestKB                 = $NewestKB
    NewestFullVersion        = $NewestFullVersion
    ApiProvider              = $ApiProvider
    ApiVersion               = $ApiVersion
    ApiDataset               = $ApiDataset
    ApiLastModified          = $ApiLastModified
    ApiLastCollected         = $ApiLastCollected
    Description              = "Public Windows update reference data used for Power BI reporting."
}

Save-JsonObject -InputObject $SourceInfo -Path "$DataPath/UpdateSources.json" -Depth 10

Write-Host "Saved UpdateSources.json"

Write-Host "Windows update reference data generation completed successfully."
