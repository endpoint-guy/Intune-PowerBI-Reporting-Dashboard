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

    [PSCustomObject]@{
        OSType       = $Update.OSType
        Product      = $Product
        MajorVersion = $Update.MajorVersion
        Version      = $Update.WindowsVersion
        KB           = $Update.KBNumber
        Build        = $Update.OSBuild
        FullVersion  = $Update.FullVersion
        BuildBranch  = $BuildBranch
        UBR          = $UBR
        ReleaseDate  = $Update.ReleaseDate
        ReleaseType  = $Update.ReleaseType
        IsExpired    = $Update.IsExpired
        ArticleUrl   = $Update.ArticleUrl
        Source       = "DataForNerds"
        SourceUrl    = $SourceUrl
        GeneratedUtc = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
    }
}

$ClientUpdates = $AllUpdates |
    Where-Object {
        $_.OSType -eq "Client" -and
        $null -ne $_.BuildBranch -and
        $null -ne $_.UBR -and
        $_.IsExpired -eq $false
    } |
    Sort-Object Product, Version, BuildBranch, UBR, ReleaseDate -Unique

$ClientUpdates |
    ConvertTo-Json -Depth 10 |
    Set-Content "$DataPath/WindowsUpdateCatalog.json" -Encoding UTF8

Write-Host "Saved WindowsUpdateCatalog.json"

$LatestQualityUpdates = $ClientUpdates |
    Where-Object {
        $_.ReleaseType -in @("Standard", "Out-of-band", "Hotpatch", "Hotpatch-OOB")
    } |
    Group-Object Product, Version, BuildBranch |
    ForEach-Object {
        $_.Group |
            Sort-Object UBR -Descending |
            Select-Object -First 1
    } |
    Sort-Object Product, Version, BuildBranch

$LatestQualityUpdates |
    ConvertTo-Json -Depth 10 |
    Set-Content "$DataPath/LatestQualityUpdates.json" -Encoding UTF8

Write-Host "Saved LatestQualityUpdates.json"

$PreviewUpdates = $ClientUpdates |
    Where-Object {
        $_.ReleaseType -eq "Preview"
    } |
    Sort-Object Product, Version, BuildBranch, UBR

$PreviewUpdates |
    ConvertTo-Json -Depth 10 |
    Set-Content "$DataPath/PreviewUpdates.json" -Encoding UTF8

Write-Host "Saved PreviewUpdates.json"

$SourceInfo = [PSCustomObject]@{
    Name         = "DataForNerds Windows Update History"
    Url          = $SourceUrl
    GeneratedUtc = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
    RecordCount  = $ClientUpdates.Count
    Description  = "Public Windows update reference data used for Power BI reporting."
}

$SourceInfo |
    ConvertTo-Json -Depth 5 |
    Set-Content "$DataPath/UpdateSources.json" -Encoding UTF8

Write-Host "Saved UpdateSources.json"

Write-Host "Windows update reference data generation completed successfully."
