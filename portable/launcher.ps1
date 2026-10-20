# Adapted from portable-python (MIT); attribution is retained in LICENSE.
[CmdletBinding()]
param([switch]$NoPause, [ValidateSet('review','change')] [string]$Tool = 'review')
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
Set-StrictMode -Version 2.0
$Mode = 'setup'
$SupportRoot = $PSScriptRoot
$Root = Split-Path -Parent $SupportRoot
$StateRoot = Join-Path $SupportRoot 'state'
$LocalTemp = Join-Path $StateRoot 'temp'
$UvCache = Join-Path $StateRoot 'cache'
$RuntimeRoot = Join-Path $StateRoot 'python'
$RuntimeBin = Join-Path $StateRoot 'python-bin'
$VenvRoot = Join-Path $StateRoot 'venv'
$VenvPython = Join-Path $VenvRoot 'Scripts\python.exe'
$Downloads = Join-Path $StateRoot 'downloads'
$Backend = Join-Path $StateRoot 'backend'
$ReadyMarker = Join-Path $StateRoot 'ready.state'
$LogPath = Join-Path $StateRoot 'setup.jsonl'
$ManifestPath = Join-Path $SupportRoot 'manifest.json'
$PinsPath = Join-Path $SupportRoot 'pins.json'
$Requirements = Join-Path $SupportRoot 'requirements.txt'
$UvExe = Join-Path $StateRoot 'uv.exe'
$LocalRootApproved = $false
$WorkspaceLock = $null
$OverallExitCode = 10
$Pins = $null

function ConvertTo-NeutralText([string]$Text) {
    foreach ($privatePath in @($Root, $env:USERPROFILE, $env:TEMP, $env:TMP)) {
        if ($privatePath) { $Text = [regex]::Replace($Text, [regex]::Escape($privatePath), '<local>', 'IgnoreCase') }
    }
    return $Text
}

function Set-PortableEnvironment {
    Get-ChildItem Env: | Where-Object { $_.Name -like 'UV_*' } | ForEach-Object {
        [Environment]::SetEnvironmentVariable($_.Name, $null, 'Process')
    }
    $env:UV_CACHE_DIR = $UvCache
    $env:UV_PROJECT_ENVIRONMENT = $VenvRoot
    $env:UV_PYTHON_INSTALL_DIR = $RuntimeRoot
    $env:UV_PYTHON_BIN_DIR = $RuntimeBin
    $env:PYTHONPATH = $null
    $env:PYTHONHOME = $null
    $env:PYTHONPYCACHEPREFIX = $null
    $env:PYTHONNOUSERSITE = '1'
    $env:PYTHONDONTWRITEBYTECODE = '1'
    $env:TEMP = $LocalTemp
    $env:TMP = $LocalTemp
}

function Get-Sha256([string]$Path) {
    $algorithm = [System.Security.Cryptography.SHA256]::Create()
    $stream = [System.IO.File]::OpenRead($Path)
    try {
        return ([System.BitConverter]::ToString($algorithm.ComputeHash($stream))).Replace("-", "").ToLowerInvariant()
    } finally {
        $stream.Dispose()
        $algorithm.Dispose()
    }
}

function Write-PortableLog(
    [string]$Phase,
    [string]$Event,
    [string]$Status,
    [string]$Detail
) {
    if (-not $LocalRootApproved) { return }
    try {
        if (-not (Test-Path -LiteralPath $StateRoot -PathType Container)) {
            New-Item -ItemType Directory -Force -Path $StateRoot | Out-Null
        }
        $record = [ordered]@{
            timestamp = [DateTime]::UtcNow.ToString("o")
            phase = $Phase
            event = $Event
            status = $Status
            detail = $Detail
        }
        if ((Test-Path -LiteralPath $LogPath) -and (Get-Item -LiteralPath $LogPath).Length -gt 8MB) { return }
        Add-Content -LiteralPath $LogPath -Value ($record | ConvertTo-Json -Compress) -Encoding UTF8
    } catch {
        # Logging is diagnostic and must not replace the real failure.
    }
}

function Confirm-LocalFixedDrive([string]$Path) {
    $fullPath = [System.IO.Path]::GetFullPath($Path)
    if ($fullPath.StartsWith("\\")) {
        throw "Copy and extract the project locally before starting it; UNC paths are not supported."
    }
    $driveRoot = [System.IO.Path]::GetPathRoot($fullPath)
    $drive = New-Object System.IO.DriveInfo($driveRoot)
    if ($drive.DriveType -ne [System.IO.DriveType]::Fixed) {
        throw "The project must run from a local fixed drive, not $driveRoot ($($drive.DriveType))."
    }
    $ancestor = Get-Item -LiteralPath $fullPath
    while ($null -ne $ancestor) {
        if ($ancestor.Attributes -band [System.IO.FileAttributes]::ReparsePoint) {
            throw "Use a real local directory, not a junction or symbolic link."
        }
        $ancestor = $ancestor.Parent
    }
}

function Invoke-Checked(
    [string]$Name,
    [string]$Executable,
    [string[]]$Arguments
) {
    Write-Host "[$Name]"
    Write-PortableLog $Mode "command_start" "running" $Name
    $timer = [System.Diagnostics.Stopwatch]::StartNew()
    $previousPreference = $ErrorActionPreference
    try {
        $ErrorActionPreference = "Continue"
        & $Executable @Arguments 2>&1 | ForEach-Object {
            $line = ConvertTo-NeutralText ([string]$_)
            Write-Host $line
        }
        $code = $LASTEXITCODE
    } finally {
        $ErrorActionPreference = $previousPreference
        $timer.Stop()
    }
    Write-PortableLog $Mode "command_finish" "info" "$Name; exit code $code; seconds=$($timer.Elapsed.TotalSeconds.ToString('F2', [System.Globalization.CultureInfo]::InvariantCulture))"
    if ($code -ne 0) {
        $exception = New-Object System.Exception("$Name failed with exit code $code")
        $exception.Data["NativeExitCode"] = $code
        throw $exception
    }
}

function Download-VerifiedFile(
    [string]$Url,
    [string]$Destination,
    [string]$ExpectedSha256
) {
    if (Test-Path -LiteralPath $Destination -PathType Leaf) {
        if ((Get-Sha256 $Destination) -eq $ExpectedSha256) {
            Write-Host "Using verified download: $([System.IO.Path]::GetFileName($Destination))"
            return
        }
        [System.IO.File]::Delete($Destination)
    }

    $partial = "$Destination.part"
    if (Test-Path -LiteralPath $partial) {
        [System.IO.File]::Delete($partial)
    }

    Write-Host "Download: $Url"
    try {
        $defaultProxy = [System.Net.WebRequest]::DefaultWebProxy
        if ($null -ne $defaultProxy) {
            $defaultProxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
        }
        Invoke-WebRequest -UseBasicParsing -Uri $Url -OutFile $partial
    } catch {
        if (Test-Path -LiteralPath $partial) {
            [System.IO.File]::Delete($partial)
        }
        Write-Host "PowerShell download failed; retrying with Windows BITS proxy authentication."
        try {
            Import-Module BitsTransfer -ErrorAction Stop
            $bitsSucceeded = $false
            foreach ($proxyUsage in @("AutoDetect", "SystemDefault")) {
                try {
                    Start-BitsTransfer -Source $Url -Destination $partial -TransferType Download -Priority Foreground -ProxyUsage $proxyUsage -ProxyAuthentication Negotiate -ErrorAction Stop
                    $bitsSucceeded = $true
                    break
                } catch {
                    if (Test-Path -LiteralPath $partial) {
                        [System.IO.File]::Delete($partial)
                    }
                }
            }
            if (-not $bitsSucceeded) {
                throw "BITS could not download through the available network path."
            }
        } catch {
            if (Test-Path -LiteralPath $partial) {
                [System.IO.File]::Delete($partial)
            }
            throw "The official component archive could not be downloaded. Check Internet/proxy access and start again."
        }
    }

    $actualSha256 = Get-Sha256 $partial
    if ($actualSha256 -ne $ExpectedSha256) {
        [System.IO.File]::Delete($partial)
        throw "Downloaded component failed pinned SHA-256 verification."
    }
    Move-Item -LiteralPath $partial -Destination $Destination -Force
}

function Install-VerifiedUv {
    $uvReady = (
        (Test-Path -LiteralPath $UvExe -PathType Leaf) -and
        ((Get-Sha256 $UvExe) -eq $UvExeSha256)
    )
    if (-not $uvReady) {
        $archive = Join-Path $StateRoot "uv-$UvVersion-windows-x86_64.zip"
        $extractRoot = Join-Path $StateRoot "uv-extract"
        Download-VerifiedFile $UvArchiveUrl $archive $UvArchiveSha256

        if (Test-Path -LiteralPath $extractRoot) {
            [System.IO.Directory]::Delete($extractRoot, $true)
        }
        Expand-Archive -LiteralPath $archive -DestinationPath $extractRoot
        $extractedUv = Join-Path $extractRoot "uv.exe"
        if (-not (Test-Path -LiteralPath $extractedUv -PathType Leaf)) {
            throw "The verified uv archive does not contain uv.exe."
        }
        if ((Get-Sha256 $extractedUv) -ne $UvExeSha256) {
            throw "Extracted uv.exe failed pinned SHA-256 verification."
        }
        Move-Item -LiteralPath $extractedUv -Destination $UvExe -Force
        [System.IO.Directory]::Delete($extractRoot, $true)
        [System.IO.File]::Delete($archive)
    }

    $versionOutput = (& $UvExe --version | Out-String).Trim()
    if ($LASTEXITCODE -ne 0 -or $versionOutput -notmatch ('^uv ' + [regex]::Escape($UvVersion) + '\b')) {
        throw "Unexpected uv executable: $versionOutput"
    }
    Write-Host "uv:            $versionOutput"
}


function Confirm-ToolFiles {
    $manifest = Get-Content -LiteralPath $ManifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
    foreach ($entry in $manifest.immutable_sha256.PSObject.Properties) {
        $relative = $entry.Name
        if ($relative -match '(^[\\/]|:|(^|[\\/])\.\.([\\/]|$))') { throw 'Invalid manifest path.' }
        $path = Join-Path $Root $relative
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw 'A tool file is missing. Extract a fresh ZIP.' }
        if ((Get-Item -LiteralPath $path).Attributes -band [IO.FileAttributes]::ReparsePoint) { throw 'Tool files must not be links.' }
        if ((Get-Sha256 $path) -ne $entry.Value) { throw 'A tool file changed. Extract a fresh ZIP; edit only the two INI files.' }
    }
}

function Confirm-PinnedRuntime {
    foreach ($entry in $Pins.python.core_sha256.PSObject.Properties) {
        $path = Join-Path (Split-Path -Parent $RuntimePython) $entry.Name
        if ((Test-Path -LiteralPath $path) -and ((Get-Item -LiteralPath $path).Attributes -band [IO.FileAttributes]::ReparsePoint)) { throw 'Runtime/backend files must not be links.' }
        if (-not (Test-Path -LiteralPath $path -PathType Leaf) -or (Get-Sha256 $path) -ne $entry.Value) {
            throw 'The local Python runtime is incomplete or changed. Extract a fresh ZIP.'
        }
    }
}

function Confirm-Backend {
    foreach ($entry in $Pins.sevenzip.sha256.PSObject.Properties) {
        $path = Join-Path $Backend $entry.Name
        if ((Test-Path -LiteralPath $path) -and ((Get-Item -LiteralPath $path).Attributes -band [IO.FileAttributes]::ReparsePoint)) { throw 'Runtime/backend files must not be links.' }
        if (-not (Test-Path -LiteralPath $path -PathType Leaf) -or (Get-Sha256 $path) -ne $entry.Value) {
            throw 'The local archive backend is incomplete or changed. Extract a fresh ZIP.'
        }
    }
}

function Get-LocationFingerprint {
    $bytes = [Text.Encoding]::UTF8.GetBytes([IO.Path]::GetFullPath($Root).ToLowerInvariant())
    $algorithm = [Security.Cryptography.SHA256]::Create()
    try { return ([BitConverter]::ToString($algorithm.ComputeHash($bytes))).Replace('-', '').ToLowerInvariant() }
    finally { $algorithm.Dispose() }
}

function Sync-LocalEnvironment([switch]$Offline) {
    $network = if ($Offline) { @('--offline') } else { @('--system-certs') }
    Invoke-Checked 'environment_create' $UvExe (@('--no-config','--managed-python','--no-python-downloads') + $network + @(
        'venv','--clear','--python',$RuntimePython,$VenvRoot
    ))
    Invoke-Checked 'packages_install' $UvExe (@('--no-config','--no-python-downloads') + $network + @(
        'pip','install','--python',$VenvPython,'--require-hashes','--only-binary',':all:',
        '--no-deps','--link-mode','copy','--index-url','https://pypi.org/simple','-r',$Requirements
    ))
}

function Confirm-Environment {
    Invoke-Checked 'environment_check' $VenvPython @('-I','-B','-X','utf8',
        (Join-Path $SupportRoot 'runtime_check.py'),$PinsPath,$RuntimePython)
}

function Save-Readiness {
    $record = [ordered]@{ schema=1; manifest_sha256=(Get-Sha256 $ManifestPath); location=(Get-LocationFingerprint) }
    [IO.File]::WriteAllText($ReadyMarker,($record | ConvertTo-Json),[Text.Encoding]::UTF8)
}

try {
    if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT -or $env:PROCESSOR_ARCHITECTURE -ne 'AMD64') {
        throw 'Use 64-bit Windows PowerShell on Windows x86-64.'
    }
    Confirm-LocalFixedDrive $Root
    Confirm-LocalFixedDrive $SupportRoot
    Confirm-ToolFiles
    foreach ($path in @($StateRoot,$LocalTemp,$UvCache,$RuntimeRoot,$RuntimeBin,$VenvRoot,$Downloads,$Backend)) {
        if (Test-Path -LiteralPath $path) { Confirm-LocalFixedDrive $path }
    }
    $LocalRootApproved = $true
    New-Item -ItemType Directory -Force -Path $StateRoot,$LocalTemp,$Downloads | Out-Null
    foreach ($path in @($ReadyMarker,$LogPath,(Join-Path $StateRoot 'workspace.lock'))) {
        if ((Test-Path -LiteralPath $path) -and ((Get-Item -LiteralPath $path).Attributes -band [IO.FileAttributes]::ReparsePoint)) {
            throw 'Generated state files must not be links.'
        }
    }
    try { $WorkspaceLock = [IO.File]::Open((Join-Path $StateRoot 'workspace.lock'),'OpenOrCreate','ReadWrite','None') }
    catch { throw 'This tool folder is already in use. Close the other scan/setup first.' }
    $Pins = Get-Content -LiteralPath $PinsPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $UvVersion = $Pins.uv.version
    $UvArchiveUrl = $Pins.uv.url
    $UvArchiveSha256 = $Pins.uv.archive_sha256
    $UvExeSha256 = $Pins.uv.exe_sha256
    if ($UvVersion -ne '0.13.1' -or $Pins.python.version -ne '3.13.16') { throw 'Unsupported toolchain pins.' }
    $PythonRequest = 'cpython-' + $Pins.python.version + '-windows-x86_64-none'
    $RuntimePython = Join-Path $RuntimeRoot ($PythonRequest + '\python.exe')
    Set-PortableEnvironment
    Set-Location -LiteralPath $Root
    Write-Host ('FOLDER ' + $(if ($Tool -eq 'change') { 'CHANGE REVIEW' } else { 'REVIEW' }))
    if (Test-Path -LiteralPath $ReadyMarker -PathType Leaf) {
        $ready = Get-Content -LiteralPath $ReadyMarker -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($ready.schema -ne 1 -or $ready.manifest_sha256 -ne (Get-Sha256 $ManifestPath)) { throw 'Prepared tool files changed. Extract a fresh ZIP.' }
        Confirm-PinnedRuntime
        Confirm-Backend
        if ($ready.location -ne (Get-LocationFingerprint)) {
            Write-Host 'Folder moved: rebuilding only the local environment from its cache (offline).'
            if (-not (Test-Path -LiteralPath $UvExe -PathType Leaf) -or (Get-Sha256 $UvExe) -ne $UvExeSha256) {
                throw 'Offline relocation needs the retained verified uv executable. Use a fresh extraction for online setup.'
            }
            Sync-LocalEnvironment -Offline
            Confirm-Environment
            Save-Readiness
        } else { Confirm-Environment }
        Write-Host 'Setup: SKIPPED (offline)'
        Write-PortableLog 'setup' 'ready_check' 'success' 'offline'
    } else {
        Write-Host 'First-time setup: downloading the pinned Python, packages and archive backend.'
        Write-PortableLog 'setup' 'setup_start' 'running' 'pinned_release'
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Install-VerifiedUv
        if (-not (Test-Path -LiteralPath $RuntimePython -PathType Leaf)) {
            $catalogOutput = & $UvExe --no-config --offline python list $PythonRequest --all-versions --all-platforms --all-arches --only-downloads --show-urls --output-format json
            if ($LASTEXITCODE -ne 0) { throw 'Cannot inspect the pinned Python catalog.' }
            $catalog = @($catalogOutput | ConvertFrom-Json)
            if ($catalog.Count -ne 1 -or ([string]$catalog[0].url).Replace('%2B','+') -ne ([string]$Pins.python.url).Replace('%2B','+')) {
                throw 'Pinned Python catalog origin mismatch.'
            }
            Invoke-Checked 'python_install' $UvExe @('--no-config','--system-certs','--managed-python','python','install',
                $PythonRequest,'--install-dir',$RuntimeRoot,'--no-bin','--no-registry')
        }
        Confirm-PinnedRuntime
        Sync-LocalEnvironment
        $installer = Join-Path $Downloads '7zip-installer.exe'
        Download-VerifiedFile $Pins.sevenzip.installer_url $installer $Pins.sevenzip.installer_sha256
        $decoder = Join-Path $Downloads '7zr.exe'
        Download-VerifiedFile $Pins.sevenzip.bootstrap.url $decoder $Pins.sevenzip.bootstrap.sha256
        New-Item -ItemType Directory -Force -Path $Backend | Out-Null
        Invoke-Checked 'backend_prepare' $decoder @('e',$installer,('-o' + $Backend),'-y','7z.exe','7z.dll','License.txt')
        Confirm-Backend
        Confirm-Environment
        Save-Readiness
        Write-PortableLog 'setup' 'setup_finish' 'success' 'pinned_release'
        Write-Host 'Setup: COMPLETE'
    }
    $Mode = 'review'
    Write-PortableLog 'review' 'run_start' 'running' 'ini_config'
    $previousPreference = $ErrorActionPreference
    try {
        $ErrorActionPreference = 'Continue'
        $entry = if ($Tool -eq 'change') { 'run_change_review.py' } else { 'run_review.py' }
        & $VenvPython -u -E -s -B -X utf8 (Join-Path $SupportRoot $entry)
        $OverallExitCode = $LASTEXITCODE
    } finally { $ErrorActionPreference = $previousPreference }
    Write-PortableLog 'review' 'run_finish' 'finished' ('exit_code=' + $OverallExitCode)
} catch {
    Write-Host ('ERROR: ' + (ConvertTo-NeutralText $_.Exception.Message)) -ForegroundColor Red
    Write-PortableLog $Mode 'failure' 'failure' 'operation_failed'
    $OverallExitCode = if ($Mode -eq 'review') { 1 } else { 10 }
} finally {
    if ($null -ne $WorkspaceLock) { $WorkspaceLock.Dispose() }
    Write-Host ('Exit code: ' + $OverallExitCode)
    if (-not $NoPause) { [void](Read-Host 'Press Enter to close') }
}
exit $OverallExitCode
