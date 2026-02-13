param(
  [switch]$SkipPubGet
)

$ErrorActionPreference = 'Stop'
$PSNativeCommandUseErrorActionPreference = $true

$projectRoot = Split-Path -Parent $PSScriptRoot
Set-Location $projectRoot

function Invoke-CheckedCommand {
  param(
    [Parameter(Mandatory = $true)][string]$Command,
    [Parameter()][string[]]$Arguments = @()
  )

  & $Command @Arguments
  if ($LASTEXITCODE -ne 0) {
    throw "Command failed: $Command $($Arguments -join ' ') (exit=$LASTEXITCODE)"
  }
}

Write-Host '[1/3] flutter pub get' -ForegroundColor Cyan
if (-not $SkipPubGet) {
  Invoke-CheckedCommand -Command 'flutter' -Arguments @('pub', 'get')
} else {
  Write-Host 'Skipped by -SkipPubGet' -ForegroundColor Yellow
}

Write-Host '[2/3] flutter analyze' -ForegroundColor Cyan
Invoke-CheckedCommand -Command 'flutter' -Arguments @('analyze', '--no-fatal-infos')

Write-Host '[3/3] flutter test' -ForegroundColor Cyan
Invoke-CheckedCommand -Command 'flutter' -Arguments @('test')

Write-Host 'All checks passed.' -ForegroundColor Green
