param([string]$SevenZipDir = "$env:ProgramFiles\7-Zip")
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
Set-Location -LiteralPath $root
if (-not [Environment]::Is64BitOperatingSystem) { throw 'Build on Windows x86-64.' }
$uv = Get-Command uv.exe -ErrorAction SilentlyContinue
if (-not $uv) {
    $fallback = Join-Path $env:USERPROFILE '.local\bin\uv.exe'
    if (Test-Path -LiteralPath $fallback) { $uv = Get-Item -LiteralPath $fallback }
}
if (-not $uv) { throw 'uv.exe is required on the build machine.' }
$uvPath = $uv.Source
if (-not $uvPath) { $uvPath = $uv.FullName }

function Invoke-Native([string]$File, [string[]]$Arguments) {
    # PowerShell 5.1 treats ordinary native stderr as ErrorRecord objects.
    # Native exit codes, rather than stderr output, determine success.
    $old = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        & $File @Arguments
        if ($LASTEXITCODE -ne 0) { throw "$File failed with exit code $LASTEXITCODE" }
    } finally { $ErrorActionPreference = $old }
}

$pins = Get-Content (Join-Path $PSScriptRoot '7zip-pins.json') -Raw | ConvertFrom-Json
foreach ($entry in $pins.sha256.PSObject.Properties) {
    $file = Join-Path $SevenZipDir $entry.Name
    $hash = [Security.Cryptography.SHA256]::Create()
    $stream = [IO.File]::OpenRead($file)
    try { $actual = ([BitConverter]::ToString($hash.ComputeHash($stream))).Replace('-', '').ToLowerInvariant() }
    finally { $stream.Dispose(); $hash.Dispose() }
    if ($actual -ne $entry.Value) { throw "7-Zip pin mismatch: $file. Expected x64 version $($pins.version)." }
}

New-Item -ItemType Directory -Force -Path '.build' | Out-Null
$env:PYTHONOPTIMIZE = '0'
# Fixed build inputs, not runtime settings. Stable epoch: 2026-01-01 UTC.
$env:PYTHONHASHSEED = '0'
$env:SOURCE_DATE_EPOCH = '1767225600'
$env:PYTHONNOUSERSITE = '1'
Remove-Item Env:PYTHONHOME, Env:PYTHONPATH -ErrorAction SilentlyContinue
Invoke-Native $uvPath @('venv', '--clear', '.build-venv', '--python', '3.13.14')
$py = Join-Path $root '.build-venv\Scripts\python.exe'
Invoke-Native $py @('-c', 'import sys; assert sys.version_info[:3] == (3,13,14) and sys.maxsize > 2**32')
Invoke-Native $uvPath @('export', '--frozen', '--no-dev', '--no-emit-project', '--output-file', '.build\runtime-requirements.txt')
Invoke-Native $uvPath @('export', '--frozen', '--only-group', 'test', '--no-emit-project', '--output-file', '.build\test-requirements.txt')
Invoke-Native $uvPath @('pip', 'install', '--python', $py, '--require-hashes', '-r', '.build\runtime-requirements.txt', '-r', '.build\test-requirements.txt', '-r', 'windows_build\requirements-build.txt')

$env:FOLDER_REVIEW_7ZIP_DIR = (Resolve-Path -LiteralPath $SevenZipDir).Path
$env:PATH = $env:FOLDER_REVIEW_7ZIP_DIR + ';' + $env:PATH
$env:PYINSTALLER_ZLIB_COMPRESSION_LEVEL = '9'
Remove-Item Env:FOLDER_REVIEW_BUILD_TESTS -ErrorAction SilentlyContinue
Invoke-Native $py @('windows_build\run_tests.py', '--results', '.build\source-tests.json')
Invoke-Native $py @('-m', 'PyInstaller', '--noconfirm', '--clean', '--distpath', '.build\candidate', '--workpath', '.build\production', 'windows_build\folder-review.spec')
$env:FOLDER_REVIEW_BUILD_TESTS = '1'
try {
    Invoke-Native $py @('-m', 'PyInstaller', '--noconfirm', '--clean', '--distpath', '.build\validation', '--workpath', '.build\test-runtime', 'windows_build\folder-review.spec')
} finally { Remove-Item Env:FOLDER_REVIEW_BUILD_TESTS -ErrorAction SilentlyContinue }
Invoke-Native (Join-Path $root '.build\validation\folder-review-tests.exe') @('--results', (Join-Path $root '.build\frozen-tests.json'))
Invoke-Native $py @('windows_build\verify.py', '--exe', '.build\candidate\folder-review.exe', '--results', '.build\verification.json')
Invoke-Native $py @('windows_build\reproduce.py', '--exe', '.build\candidate\folder-review.exe', '--results', '.build\reproducibility.json')
# Publish only after all mandatory checks and release preparation succeed.
New-Item -ItemType Directory -Force -Path '.build\release' | Out-Null
Copy-Item '.build\candidate\folder-review.exe' '.build\release\folder-review.exe' -Force
Copy-Item '.build\verification.json' '.build\release\verification.json' -Force
Copy-Item '.build\source-tests.json', '.build\frozen-tests.json', '.build\reproducibility.json' '.build\release' -Force
Copy-Item 'LICENSE', 'windows_build\THIRD-PARTY-NOTICES.txt', 'windows_build\USAGE.txt' '.build\release' -Force
Copy-Item (Join-Path $SevenZipDir 'License.txt') '.build\release\7zip-license.txt' -Force
# Both deliveries come from these same inputs; manifest preparation verifies
# that no older standalone Python scanner can be published beside this EXE.
New-Item -ItemType Directory -Force -Path '.build\release\python' | Out-Null
Copy-Item 'folder_review.py', 'folder_review_ui.py', 'pyproject.toml', 'uv.lock', 'RUN.bat', 'SETUP.bat', 'README.md', 'LICENSE' '.build\release\python' -Force
Invoke-Native $py @('windows_build\manifest.py', '--exe', '.build\release\folder-review.exe', '--output', '.build\release\build-manifest.json')
New-Item -ItemType Directory -Force -Path 'dist' | Out-Null
Copy-Item '.build\release\*' 'dist' -Recurse -Force
Get-Item 'dist\folder-review.exe' | Select-Object FullName,Length
