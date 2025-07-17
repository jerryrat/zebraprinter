# Access Database Monitor Build Script
# Version: 1.0.0

param(
    [string]$Configuration = "Release",
    [string]$OutputPath = "../Release"
)

Write-Host "Starting build process..." -ForegroundColor Green

# Clean previous build
if (Test-Path $OutputPath) {
    Write-Host "Cleaning previous build..." -ForegroundColor Yellow
    Remove-Item -Path $OutputPath -Recurse -Force
}

# Create output directory
New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null

# Build and publish
Write-Host "Building and publishing application..." -ForegroundColor Yellow
dotnet publish -c $Configuration -r win-x64 --self-contained true -p:PublishSingleFile=true -p:PublishReadyToRun=true -o $OutputPath

if ($LASTEXITCODE -eq 0) {
    Write-Host "Build completed successfully!" -ForegroundColor Green
    Write-Host "Output location: $OutputPath" -ForegroundColor Cyan
    
    # List generated files
    Write-Host "`nGenerated files:" -ForegroundColor Cyan
    Get-ChildItem -Path $OutputPath | ForEach-Object {
        Write-Host "  $($_.Name)" -ForegroundColor White
    }
} else {
    Write-Host "Build failed!" -ForegroundColor Red
    exit 1
}

Write-Host "`nBuild process completed." -ForegroundColor Green 