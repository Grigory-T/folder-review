param(
    [Parameter(Mandatory=$true)][string]$Zip,
    [string]$Results
)
$ErrorActionPreference='Stop'
$utf8=New-Object Text.UTF8Encoding($false)
$base=Join-Path (Split-Path ([IO.Path]::GetFullPath($Zip)) -Parent) ('acceptance-'+[Guid]::NewGuid().ToString('N'))
$root=Join-Path $base ('tool '+[char]0x00e9)
$inputDir=Join-Path $base ('input '+[char]0x044f+' 100%')
$outside=Join-Path $base 'outside-canary'
New-Item -ItemType Directory -Force -Path $root,$inputDir,$outside | Out-Null
Expand-Archive -LiteralPath $Zip -DestinationPath $root
[IO.File]::WriteAllText((Join-Path $inputDir 'one.txt'),'same contents',$utf8)
[IO.File]::WriteAllText((Join-Path $inputDir 'two.txt'),'same contents',$utf8)
[IO.File]::WriteAllText((Join-Path $inputDir 'unique.txt'),'one different length',$utf8)
$before=@{}
Get-ChildItem -LiteralPath $inputDir -File | ForEach-Object {
    $before[$_.Name]=((Get-FileHash -LiteralPath $_.FullName).Hash+':'+$_.LastWriteTimeUtc.Ticks)
}
foreach ($name in @('TEMP','TMP','LOCALAPPDATA','APPDATA','PYTHONPYCACHEPREFIX')) {
    [Environment]::SetEnvironmentVariable($name,$outside,'Process')
}
# A malicious inherited uv override must never select another package index.
$env:UV_INDEX='https://invalid.invalid/simple'
$env:UV_PYTHON_INSTALL_DIR=$outside
$records=New-Object Collections.Generic.List[object]

function Set-Setting([string]$Name,[string]$Value,[switch]$Unicode) {
    $path=Join-Path $root 'settings.ini'
    $text=[IO.File]::ReadAllText($path)
    $evaluator=[Text.RegularExpressions.MatchEvaluator]{param($m) $Name+' = '+$Value}
    $text=[regex]::Replace($text,('(?m)^'+[regex]::Escape($Name)+'\s*=.*$'),$evaluator)
    $encoding=if($Unicode){[Text.Encoding]::Unicode}else{$utf8}
    [IO.File]::WriteAllText($path,$text,$encoding)
}

function Invoke-Tool([string]$Case,[int]$Expected,[switch]$Batch) {
    $timer=[Diagnostics.Stopwatch]::StartNew()
    $old=$ErrorActionPreference
    try {
        $ErrorActionPreference='Continue'
        if($Batch) {
            $lines=& cmd.exe /d /c ('call "'+(Join-Path $root 'SETUP.bat')+'" --no-pause') 2>&1
        } else {
            $lines=& powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File (Join-Path $root 'support\launcher.ps1') -NoPause 2>&1
        }
        $code=$LASTEXITCODE
    } finally { $ErrorActionPreference=$old; $timer.Stop() }
    $script:lastOutput=($lines | Out-String)
    if($code -ne $Expected) { throw "$Case expected $Expected, got $code`n$lastOutput" }
    if(@(Get-ChildItem -LiteralPath $outside -Force).Count -ne 0) { throw 'Tool wrote into external profile/temp canary' }
    if(@(Get-ChildItem -LiteralPath (Join-Path $root 'support\state') -Filter 'work-*' -ErrorAction SilentlyContinue).Count) { throw 'Scan workspaces leaked' }
    $records.Add([ordered]@{case=$Case;exit_code=$code;seconds=[math]::Round($timer.Elapsed.TotalSeconds,3)})
    Write-Host ($Case+': PASS')
}

Set-Setting 'folder' ('"'+$inputDir+'"') -Unicode
Invoke-Tool 'fresh-setup-and-scan' 0
if($lastOutput -notmatch 'Setup: COMPLETE' -or $lastOutput -notmatch 'Scan status: COMPLETE') { throw 'Fresh preparation did not run' }
$report=Join-Path $root 'folder-review.xlsx'
if(-not(Test-Path -LiteralPath $report)) { throw 'Report not beside the launcher' }
$ready=Join-Path $root 'support\state\ready.state'
$readyBytes=[Convert]::ToBase64String([IO.File]::ReadAllBytes($ready))
$readyTime=(Get-Item -LiteralPath $ready).LastWriteTimeUtc.Ticks
$uv=Join-Path $root 'support\state\uv.exe'
$uvSaved=$uv+'.saved'
# Prove recurring scans do not execute/download uv at all, even if it is absent.
Move-Item -LiteralPath $uv -Destination $uvSaved
$env:HTTP_PROXY='http://127.0.0.1:9'
$env:HTTPS_PROXY='http://127.0.0.1:9'
$env:ALL_PROXY='http://127.0.0.1:9'
$env:NO_PROXY=''
try {
    foreach($mode in @('off','filtered','all')) {
        Set-Setting 'duplicates' $mode
        Invoke-Tool ('offline-'+$mode) 0
        if($lastOutput -notmatch 'Setup: SKIPPED' -or $lastOutput -match 'Download:|First-time setup') { throw 'Recurring scan performed preparation' }
    }
    Invoke-Tool 'actual-batch-wrapper' 0 -Batch
} finally { Move-Item -LiteralPath $uvSaved -Destination $uv }
if([Convert]::ToBase64String([IO.File]::ReadAllBytes($ready)) -ne $readyBytes -or (Get-Item -LiteralPath $ready).LastWriteTimeUtc.Ticks -ne $readyTime) { throw 'Unchanged starts rewrote readiness' }
Set-Setting 'workers' '0'
$reportBefore=(Get-FileHash -LiteralPath $report).Hash
Invoke-Tool 'invalid-settings-preserve-report' 1
if((Get-FileHash -LiteralPath $report).Hash -ne $reportBefore) { throw 'Invalid settings replaced the last report' }
Set-Setting 'workers' 'auto'
$stream=[IO.File]::Open((Join-Path $root 'support\state\workspace.lock'),'OpenOrCreate','ReadWrite','None')
try { Invoke-Tool 'concurrent-workspace-rejected' 10 } finally { $stream.Dispose() }
$core=Join-Path $root 'support\folder_review.py'
$original=[IO.File]::ReadAllBytes($core)
[IO.File]::AppendAllText($core,"`n# changed",$utf8)
try { Invoke-Tool 'changed-technical-source-rejected' 10 } finally { [IO.File]::WriteAllBytes($core,$original) }

# Archive coverage failure remains a usable PARTIAL / exit 2, not a setup error.
$partial=Join-Path $base 'partial-input'
New-Item -ItemType Directory -Path $partial | Out-Null
Copy-Item (Join-Path $PSScriptRoot 'fixtures\windows\rar5-encrypted-headers.rar') (Join-Path $partial 'encrypted.rar')
Set-Setting 'folder' $partial
Invoke-Tool 'encrypted-header-partial' 2
$unc='\\localhost\C$\'+$inputDir.Substring(3)
$smbAvailable=Test-Path -LiteralPath $unc -ErrorAction SilentlyContinue
if($smbAvailable) {
    Set-Setting 'folder' $unc
    Invoke-Tool 'SMB-input-local-report' 0
}
Set-Setting 'folder' $root
Invoke-Tool 'self-scan' 0
Set-Setting 'folder' $inputDir

# Copy all state; recreate only the local environment from copied cache, offline.
$previous=$root
$root=Join-Path $base 'relocated tool'
Copy-Item -LiteralPath $previous -Destination $root -Recurse
Invoke-Tool 'offline-relocation-and-scan' 0
if($lastOutput -notmatch 'Folder moved:') { throw 'Relocation did not rebuild venv' }
$previous=$root
$root=Join-Path $base 'relocated-without-uv'
Copy-Item -LiteralPath $previous -Destination $root -Recurse
Remove-Item -LiteralPath (Join-Path $root 'support\state\uv.exe')
Invoke-Tool 'incomplete-offline-relocation-fails-closed' 10
if($lastOutput -match 'Download:|First-time setup') { throw 'Relocation attempted online fallback' }
$root=$previous

foreach($path in Get-ChildItem -LiteralPath $inputDir -File) {
    if($before[$path.Name] -ne ((Get-FileHash -LiteralPath $path.FullName).Hash+':'+$path.LastWriteTimeUtc.Ticks)) { throw 'Input contents or modification times changed' }
}
$journal=[IO.File]::ReadAllText((Join-Path $root 'support\state\setup.jsonl'))
foreach($token in @($inputDir,$root,$env:USERNAME,'one.txt','same contents')) {
    if($token -and $journal.IndexOf($token,[StringComparison]::OrdinalIgnoreCase) -ge 0) { throw 'Private input/identity in setup journal' }
}
$receipt=[ordered]@{
    zip_sha256=(Get-FileHash -LiteralPath $Zip).Hash.ToLowerInvariant()
    all_checks_passed=$true
    input_contents_and_modification_times_preserved=$true
    external_profile_temp_canary_untouched=$true
    neutral_setup_journal=$true
    fixed_source_integrity_checked=$true
    smb_input_checked=$smbAvailable
    cases=$records.ToArray()
}
if($Results) { [IO.File]::WriteAllText($Results,($receipt|ConvertTo-Json -Depth 6),$utf8) }
Write-Host 'WINDOWS PORTABLE ACCEPTANCE: SUCCESS'

exit 0
