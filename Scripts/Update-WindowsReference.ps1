$ErrorActionPreference = "Stop"

$DataPath = "./Data"

if (-not (Test-Path $DataPath)) {
    New-Item -Path $DataPath -ItemType Directory -Force | Out-Null
}

$SourceUrl = "https://api.datafornerds.io/v2/microsoft/windows-update-history.json"

Write-Host "Downloading Windows Update history from public API..."
Write-Host "Source: $SourceUrl"

$Response = Invoke-RestMethod -Uri $SourceUrl -Method Get

if ($null -eq $Response.data) {
    throw "No data property found in API response."
}

Write-Host "Records received from API: $($Response.data.Count)"

$AllUpdates = foreach ($Update in $Response.data) {
    $BuildParts = $Update.OSBuild -split "\."

    if ($BuildParts.Count -eq 2) {
        $BuildBranch = [int]$BuildParts[0]
        $UBR = [int]$BuildParts[1]
    }
    else {
        $BuildBranch = $null
        $UBR = $null
    }

    if ($Update.MajorVersion -eq 11) {
        $Product = "Windows 11"
    }
    elseif ($Update.MajorVersion -eq 10) {
        $Product = "Windows 10"
    }
    else {
        $Product = "Windows Server"
    }

    $ReleaseDateValue = $null

    if (-not :IsNullOrWhiteSpace($Update.ReleaseDate)) {
        try {
            $ReleaseDateValue = [datetime]$Update.ReleaseDate
        }
        catch {
            $ReleaseDateValue = $null
        }
    }

    [PSCustomObject]@{
        OSType          = $Update.OSType
        Product         = $Product
        MajorVersion    = $Update.MajorVersion
        Version         = $Update.WindowsVersion
        KB              = $Update.KBNumber
        Build           = $Update.OSBuild
        FullVersion     = $Update.FullVersion
        BuildBranch     = $BuildBranch
        UBR             = $UBR
        ReleaseDate     = if ($ReleaseDateValue) { $ReleaseDateValue.ToString("yyyy-MM-dd") } else { $null }
        ReleaseDateSort = $ReleaseDateValue
        ReleaseType     = $Update.ReleaseType
        IsExpired       = $Update.IsExpired
        ArticleUrl      = $Update.ArticleUrl
        Source          = "DataForNerds"
        SourceUrl       = $SourceUrl
        GeneratedUtc    = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
    }
}

$ClientUpdates = $AllUpdates |
    Where-Object {
        $_.OSType -eq "Client" -and
        $null -ne $_.BuildBranch -and
        $null -ne $_.UBR -and
        $null -ne $_.ReleaseDateSort
    } |
    Sort-Object Product, Version, BuildBranch, UBR, ReleaseDateSort -Unique

Write-Host "Client update records after filtering: $($ClientUpdates.Count)"

$NewestClientUpdate = $ClientUpdates |
    Sort-Object ReleaseDateSort -Descending |
    Select-Object -First 1

if ($NewestClientUpdate) {
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

$ExportClientUpdates = $ClientUpdates |
    Select-Object `
        OSType,
        Product,
        MajorVersion,
        Version,
        KB,
        Build,
        FullVersion,
        BuildBranch,
        UBR,
        ReleaseDate,
        ReleaseType,
        IsExpired,
        ArticleUrl,
        Source,
        SourceUrl,
        GeneratedUtc

$ExportClientUpdates |
    ConvertTo-Json -Depth 10 |
    Set-Content "$DataPath/WindowsUpdateCatalog.json" -Encoding UTF8

Write-Host "Saved WindowsUpdateCatalog.json"

$LatestQualityUpdates = $ClientUpdates |
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
    Sort-Object Product, Version, BuildBranch

Write-Host "Latest quality update records generated: $($LatestQualityUpdates.Count)"

$ExportLatestQualityUpdates = $LatestQualityUpdates |
    Select-Object `
        OSType,
        Product,
        MajorVersion,
        Version,
        KB,
        Build,
        FullVersion,
        BuildBranch,
        UBR,
        ReleaseDate,
        ReleaseType,
        IsExpired,
        ArticleUrl,
        Source,
        SourceUrl,
        GeneratedUtc

$ExportLatestQualityUpdates |
    ConvertTo-Json -Depth 10 |
    Set-Content "$DataPath/LatestQualityUpdates.json" -Encoding UTF8

Write-Host "Saved LatestQualityUpdates.json"

$PreviewUpdates = $ClientUpdates |
    Where-Object {
        $_.ReleaseType -match "Preview"
    } |
    Sort-Object Product, Version, BuildBranch, ReleaseDateSort, UBR

$ExportPreviewUpdates = $PreviewUpdates |
    Select-Object `
        OSType,
        Product,
        MajorVersion,
        Version,
        KB,
        Build,
        FullVersion,
        BuildBranch,
        UBR,
        ReleaseDate,
        ReleaseType,
        IsExpired,
        ArticleUrl,
        Source,
        SourceUrl,
        GeneratedUtc

$ExportPreviewUpdates |
    ConvertTo-Json -Depth 10 |
    Set-Content "$DataPath/PreviewUpdates.json" -Encoding UTF8

Write-Host "Saved PreviewUpdates.json"

$SourceInfo = [PSCustomObject]@{
    Name                   = "DataForNerds Windows Update History"
    Url                    = $SourceUrl
    GeneratedUtc           = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
    TotalApiRecords         = $Response.data.Count
    ClientRecordCount       = $ClientUpdates.Count
    LatestQualityRecordCount = $LatestQualityUpdates.Count
    NewestReleaseDate       = if ($NewestClientUpdate) { $NewestClientUpdate.ReleaseDate } else { $null }
    NewestKB                = if ($NewestClientUpdate) { $NewestClientUpdate.KB } else { $null }
    NewestFullVersion       = if ($NewestClientUpdate) { $NewestClientUpdate.FullVersion } else { $null }
    Description             = "Public Windows update reference data used for Power BI reporting."
}

$SourceInfo |
    ConvertTo-Json -Depth 5 |
    Set-Content "$DataPath/UpdateSources.json" -Encoding UTF8

Write-Host "Saved UpdateSources.json"

Write-Host "Windows update reference data generation completed successfully."
