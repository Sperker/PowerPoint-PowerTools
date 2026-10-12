<#
.SYNOPSIS
    Builds and publishes all PowerPoint PowerTools Add-Ins with ClickOnce packages,
    ensuring standardized documentation and updated ZIP files.
#>
param(
    [string]$TargetAddIn = "All"
)

$ErrorActionPreference = "Stop"
$root = $PSScriptRoot
$msbuild = "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe"

if (-not (Test-Path $msbuild)) {
    throw "MSBuild not found at: $msbuild"
}

$addIns = @(
    @{
        Name = "PowerPointCornerRadiusPointsAddIn"
        Project = "src\PowerPointCornerRadiusPointsAddIn\PowerPointCornerRadiusPointsAddIn.csproj"
        ReleaseDir = "releases\PowerPointCornerRadiusPointsAddIn_V1.0_RELEASE"
        ZipName = "PowerPointCornerRadiusPointsAddIn_V1.0_RELEASE.zip"
    },
    @{
        Name = "PowerPointJumpPointsAddIn"
        Project = "src\PowerPointJumpPointsAddIn\PowerPointJumpPointsAddIn.csproj"
        ReleaseDir = "releases\PowerPointJumpPointsAddIn_V2.0_RELEASE"
        ZipName = "PowerPointJumpPointsAddIn_V2.0_RELEASE.zip"
    },
    @{
        Name = "PowerPointMissingEdgeAlignAddIn"
        Project = "src\PowerPointMissingEdgeAlignAddIn\PowerPointMissingEdgeAlignAddIn.csproj"
        ReleaseDir = "releases\PowerPointMissingEdgeAlignAddIn_V1.2_RELEASE"
        ZipName = "PowerPointMissingEdgeAlignAddIn_V1.2_RELEASE.zip"
    },
    @{
        Name = "PowerPointPanZoomAddIn"
        Project = "src\PowerPointPanZoomAddIn\PowerPointPanZoomAddIn.csproj"
        ReleaseDir = "releases\PowerPointPanZoomAddIn_V1.8_RELEASE"
        ZipName = "PowerPointPanZoomAddIn_V1.8_RELEASE.zip"
    },
    @{
        Name = "PowerPointStayOnSlidePointsAddIn"
        Project = "src\PowerPointStayOnSlidePointsAddIn\PowerPointStayOnSlidePointsAddIn.csproj"
        ReleaseDir = "releases\PowerPointStayOnSlidePointsAddIn_V1.1_RELEASE"
        ZipName = "PowerPointStayOnSlidePointsAddIn_V1.1_RELEASE.zip"
    }
)

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host " PowerPoint PowerTools - Release & Publish Automation" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

foreach ($item in $addIns) {
    if ($TargetAddIn -ne "All" -and $item.Name -notlike "*$TargetAddIn*") {
        continue
    }

    $projectPath = Join-Path $root $item.Project
    $projectDir = Split-Path $projectPath -Parent
    $relDir = Join-Path $root $item.ReleaseDir
    $zipPath = Join-Path $root ("releases\" + $item.ZipName)

    Write-Host "`n--> Publishing $($item.Name)..." -ForegroundColor Yellow

    # Ensure release directory exists
    if (-not (Test-Path $relDir)) {
        New-Item -ItemType Directory -Path $relDir -Force | Out-Null
    }

    # Clean & Publish via MSBuild
    & $msbuild $projectPath /t:Clean,Build /p:Configuration=Release /v:minimal
    if ($LASTEXITCODE -ne 0) { throw "Build failed for $($item.Name)" }

    & $msbuild $projectPath /t:Publish /p:Configuration=Release /p:PublishDir="$relDir\\" /v:minimal
    if ($LASTEXITCODE -ne 0) { throw "Publish failed for $($item.Name)" }

    # Standardize documents: copy all txt and cmd from source to release folder
    Get-ChildItem -Path $projectDir -File | Where-Object { $_.Extension -in @(".txt", ".cmd") } | ForEach-Object {
        Copy-Item $_.FullName -Destination $relDir -Force
    }

    # Clean intermediate dev registry entries
    Remove-Item "HKCU:\Software\Microsoft\Office\PowerPoint\Addins\$($item.Name)" -Force -ErrorAction SilentlyContinue

    # Re-create ZIP archive
    if (Test-Path $zipPath) { Remove-Item $zipPath -Force }
    Compress-Archive -Path "$relDir\*" -DestinationPath $zipPath -Force -CompressionLevel Optimal

    Write-Host "  [OK] Released to: $relDir" -ForegroundColor Green
    Write-Host "  [OK] Packaged: $zipPath" -ForegroundColor Green
}

# Clean any leftover debug keys
Remove-Item "HKCU:\Software\Microsoft\Office\PowerPoint\Addins\PPTMissingEdgeAlignAddIn" -Force -ErrorAction SilentlyContinue

Write-Host "`nAll requested releases published and synchronized successfully!" -ForegroundColor Green
