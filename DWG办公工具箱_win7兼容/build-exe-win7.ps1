$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSCommandPath
$project = Join-Path $root 'DwgOfficeToolbox.App\DwgOfficeToolbox.App.csproj'
$distDir = Join-Path $root 'dist'
$appOutDir = Join-Path $distDir 'app'
$innoScript = Join-Path $root 'Packaging\installer-win7.iss'
$iscc = 'C:\Program Files (x86)\Inno Setup 6\ISCC.exe'

if (Test-Path $distDir) { Remove-Item $distDir -Recurse -Force }
New-Item -ItemType Directory -Path $appOutDir | Out-Null

Write-Host "Building (net48)..." -ForegroundColor Cyan
dotnet build $project -c Release

$buildOut = Join-Path $root 'DwgOfficeToolbox.App\bin\Release\net48'
Copy-Item -Recurse -Force (Join-Path $buildOut '*') $appOutDir

$odaRoot = Join-Path (Split-Path -Parent $root) 'ODAFileConverter'
if (Test-Path $odaRoot) {
    Write-Host "Copying ODAFileConverter..." -ForegroundColor Cyan
    Copy-Item -Recurse -Force $odaRoot (Join-Path $appOutDir 'ODAFileConverter')
}

if (-not (Test-Path $iscc)) {
    Write-Warning "Inno Setup not found at $iscc"
    Write-Host "Install Inno Setup 6, then run:" -ForegroundColor Yellow
    Write-Host "  `"$iscc`" `"$innoScript`"" -ForegroundColor Yellow
    exit 0
}

Write-Host "Building installer..." -ForegroundColor Cyan
& $iscc $innoScript | Out-Null

Write-Host "EXE installer created under dist\\setup" -ForegroundColor Green
