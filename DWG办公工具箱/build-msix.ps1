$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSCommandPath
$project = Join-Path $root 'DwgOfficeToolbox.App\DwgOfficeToolbox.App.csproj'
$publishDir = Join-Path $root 'msix\publish'
$layoutDir = Join-Path $root 'msix\layout'
$assetsDir = Join-Path $root 'Packaging\Assets'
$manifestPath = Join-Path $root 'Packaging\AppxManifest.xml'
$msixPath = Join-Path $root 'msix\DwgOfficeToolbox.msix'
$certDir = Join-Path $root 'msix\cert'
$pfxPath = Join-Path $certDir 'DwgOfficeToolbox.pfx'
$pfxPassword = 'DwgOfficeToolbox'

Write-Host "Publishing app..." -ForegroundColor Cyan
if (Test-Path (Join-Path $root 'DwgOfficeToolbox.App\\bin')) { Remove-Item (Join-Path $root 'DwgOfficeToolbox.App\\bin') -Recurse -Force }
if (Test-Path (Join-Path $root 'DwgOfficeToolbox.App\\obj')) { Remove-Item (Join-Path $root 'DwgOfficeToolbox.App\\obj') -Recurse -Force }
if (Test-Path $publishDir) { Remove-Item $publishDir -Recurse -Force }
dotnet publish $project -c Release -r win-x64 --self-contained true -o $publishDir

Write-Host "Preparing layout..." -ForegroundColor Cyan
if (Test-Path $layoutDir) { Remove-Item $layoutDir -Recurse -Force }
New-Item -ItemType Directory -Path $layoutDir | Out-Null
Copy-Item -Recurse -Force (Join-Path $publishDir '*') $layoutDir
Copy-Item -Force $manifestPath (Join-Path $layoutDir 'AppxManifest.xml')
Copy-Item -Recurse -Force $assetsDir (Join-Path $layoutDir 'Assets')

$kitsBin = "C:\\Program Files (x86)\\Windows Kits\\10\\bin"
$makeappxPath = ''
$signtoolPath = ''
$knownVer = "10.0.26100.0"
$knownMakeAppx = Join-Path $kitsBin ($knownVer + "\\x64\\makeappx.exe")
$knownSignTool = Join-Path $kitsBin ($knownVer + "\\x64\\signtool.exe")
if (Test-Path $knownMakeAppx) { $makeappxPath = $knownMakeAppx }
if (Test-Path $knownSignTool) { $signtoolPath = $knownSignTool }
$verDir = Get-ChildItem $kitsBin -Directory -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -match '^10\\.0\\.' } |
    Sort-Object Name -Descending |
    Select-Object -First 1

if ($verDir) {
    $makeappxPath = Join-Path $verDir.FullName 'x64\\makeappx.exe'
    $signtoolPath = Join-Path $verDir.FullName 'x64\\signtool.exe'
}

if ([string]::IsNullOrWhiteSpace($makeappxPath) -or -not (Test-Path $makeappxPath)) {
    $makeappx = Get-ChildItem $kitsBin -Recurse -Filter makeappx.exe -ErrorAction SilentlyContinue |
        Where-Object { $_.FullName -match '\\\\x64\\\\' } |
        Sort-Object FullName -Descending |
        Select-Object -First 1
    if ($makeappx) { $makeappxPath = $makeappx.FullName }
}

if ([string]::IsNullOrWhiteSpace($signtoolPath) -or -not (Test-Path $signtoolPath)) {
    $signtool = Get-ChildItem $kitsBin -Recurse -Filter signtool.exe -ErrorAction SilentlyContinue |
        Where-Object { $_.FullName -match '\\\\x64\\\\' } |
        Sort-Object FullName -Descending |
        Select-Object -First 1
    if ($signtool) { $signtoolPath = $signtool.FullName }
}

if ([string]::IsNullOrWhiteSpace($makeappxPath) -or -not (Test-Path $makeappxPath)) {
    throw "Windows SDK makeappx.exe not found."
}
if ([string]::IsNullOrWhiteSpace($signtoolPath) -or -not (Test-Path $signtoolPath)) {
    throw "Windows SDK signtool.exe not found."
}

Write-Host "Creating MSIX..." -ForegroundColor Cyan
New-Item -ItemType Directory -Path (Split-Path -Parent $msixPath) -Force | Out-Null
& $makeappxPath pack /d $layoutDir /p $msixPath /o | Out-Null

Write-Host "Signing MSIX..." -ForegroundColor Cyan
New-Item -ItemType Directory -Path $certDir -Force | Out-Null
if (-not (Test-Path $pfxPath)) {
    $cert = New-SelfSignedCertificate -Type CodeSigningCert -Subject "CN=DWGOfficeToolbox" -CertStoreLocation "Cert:\\CurrentUser\\My"
    $securePwd = ConvertTo-SecureString -String $pfxPassword -Force -AsPlainText
    Export-PfxCertificate -Cert $cert -FilePath $pfxPath -Password $securePwd | Out-Null
}

& $signtoolPath sign /fd SHA256 /a /f $pfxPath /p $pfxPassword $msixPath | Out-Null

Write-Host "MSIX created: $msixPath" -ForegroundColor Green
Write-Host "Certificate: $pfxPath (password: $pfxPassword)" -ForegroundColor Yellow
Write-Host "Import the certificate into Trusted People, then install the MSIX." -ForegroundColor Yellow
