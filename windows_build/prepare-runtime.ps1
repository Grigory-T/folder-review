$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$pins = Get-Content (Join-Path $PSScriptRoot 'python-pins.json') -Raw | ConvertFrom-Json
$downloads = Join-Path $root '.build\downloads'
$runtime = Join-Path $root '.build\python-runtime'
New-Item -ItemType Directory -Force $downloads | Out-Null
$archive = Join-Path $downloads $pins.archive
if (-not (Test-Path -LiteralPath $archive)) {
    Invoke-WebRequest -UseBasicParsing -Uri $pins.url -OutFile $archive
}
if ((Get-FileHash -LiteralPath $archive -Algorithm SHA256).Hash.ToLowerInvariant() -ne $pins.sha256) {
    throw 'Python archive hash mismatch. Build stopped.'
}
if (-not (Test-Path -LiteralPath $runtime)) {
    New-Item -ItemType Directory $runtime | Out-Null
    & tar.exe -xf $archive -C $runtime
    if ($LASTEXITCODE -ne 0) { throw 'Pinned Python extraction failed.' }
}
# Check the executable and interpreter DLL before launching either.
foreach ($entry in $pins.core_sha256.PSObject.Properties) {
    $file = Join-Path $runtime ('python\' + $entry.Name)
    if ((Get-FileHash -LiteralPath $file -Algorithm SHA256).Hash.ToLowerInvariant() -ne $entry.Value) {
        throw "Pinned Python file mismatch: $($entry.Name)"
    }
}
