$ErrorActionPreference = "Stop"

Push-Location $PSScriptRoot
try {
  if (!(Test-Path ".\\node_modules")) {
    Write-Output "Installing build dependencies (npm install)..."
    npm install | Out-Host
  }
  Write-Output "Building (.eext) via npm run build..."
  npm run build | Out-Host

  $ext = Get-Content -Raw -Encoding UTF8 ".\\extension.json" | ConvertFrom-Json
  $name = $ext.name
  $version = $ext.version
  $outFile = Join-Path $PSScriptRoot ("build\\dist\\{0}_v{1}.eext" -f $name, $version)
  Write-Output ("Built: {0}" -f $outFile)
} finally {
  Pop-Location
}
