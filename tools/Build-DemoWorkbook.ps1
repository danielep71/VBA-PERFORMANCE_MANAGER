<#
.SYNOPSIS
    Builds the demo workbook from the checked-out commit, smoke-tests every
    demo action, removes the build machine's metadata and packages it with
    LICENSE and NOTICE (#47, #72).

.DESCRIPTION
    Supervised Windows desktop Excel tool, started by Build-DemoWorkbook.cmd.
    From a clean checkout it:

      1. imports the six managed components into a new workbook, in the
         documented order, and compiles the project;
      2. runs cPM_CreateDemoSheets, which builds Home, DEMO_cPM, DATA_cPM and
         HELP_cPM, and deletes every other sheet;
      3. runs every demo action unattended (no message boxes) and checks that
         each adds a results-log row and that none fails;
      4. clears the results log and saves PERFORMANCE MANAGER.xlsm;
      5. removes the metadata a managed machine can add (sensitivity label,
         document category or content status, the folder it was saved from,
         print headers and footers) and checks that none is left;
      6. reopens the workbook with macros disabled and checks its sheets;
      7. writes VBA-Performance-Manager-v<VERSION>.zip with the workbook and
         the checkout's LICENSE and NOTICE, and a JSON build record.

    Output goes to a new folder under -OutDir (default ..\cpm-demo-build),
    outside the checkout. The checkout is never changed. AccessVBOM is enabled
    only for the run and restored on every path. Invisible Excel instances are
    closed first; a visible Excel window stops the build.
#>
[CmdletBinding()]
param(
    [string] $OutDir,
    [switch] $AllowVbomAccess
)
Set-StrictMode -Version 2
$ErrorActionPreference = 'Stop'

$repo = Split-Path -Parent $PSScriptRoot
# The documented import order, plus the demo and example modules.
$Components = @(
    'src\modules\M_cPM_TIMEWASTERS.bas',
    'src\classes\cPerformanceManager.cls',
    'demo\M_DEMO_BUILDER.bas',
    'demo\M_cPM_DEMO.bas',
    'demo\M_cPM_USAGE_EXAMPLES.bas',
    'test\M_cPM_Test.bas'
)
$BuildEntry = 'cPM_CreateDemoSheets'
$Sheets = @('Home', 'DEMO_cPM', 'DATA_cPM', 'HELP_cPM')
$LogSheet = 'DEMO_cPM'
$LogTable = 'Tbl_cPM_Log'
$SmokeActions = @('cPM_RunBasicTiming', 'cPM_RunAllMethods', 'cPM_RunAlignedDemo',
                  'cPM_RunDiagnostics', 'cPM_RunOverheadDemo', 'cPM_RunPauseDemo',
                  'cPM_RunTWComparison', 'cPM_RunMeasuredBenchmark')
$ClearAction = 'cPM_ClearResultsLog'
$WorkbookName = 'PERFORMANCE MANAGER.xlsm'
$CompileControlId = 578         # Debug > Compile VBAProject
$ExitWaitSeconds = 15
$VbomKey = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security'

if (-not $OutDir) { $OutDir = Join-Path (Split-Path -Parent $repo) 'cpm-demo-build' }
$OutDir = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutDir)
if (($OutDir.TrimEnd('\') + '\').StartsWith($repo.TrimEnd('\') + '\', [StringComparison]::OrdinalIgnoreCase)) {
    throw '-OutDir must be outside the repository.'
}

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class CpmBuildWin32 {
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
'@

function Release-Com($Object) {
    if ($null -ne $Object -and [Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)
    }
}

function Stop-WindowlessExcel {
    foreach ($process in @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue)) {
        if ($process.MainWindowTitle) {
            throw "Excel is open ('$($process.MainWindowTitle)'). Save your work, close Excel, then run again."
        }
        Write-Host "Closing an invisible Excel instance (PID $($process.Id))."
        Stop-Process -Id $process.Id -Force
    }
    Start-Sleep -Milliseconds 500
}

function Close-Excel($Excel, [uint32] $ExcelPid) {
    # Quit, release, and wait for this script's own Excel process only.
    if ($null -ne $Excel) {
        try { foreach ($book in @($Excel.Workbooks)) { $book.Close($false); Release-Com $book } } catch { }
        try { $Excel.Quit() } catch { }
        Release-Com $Excel
    }
    for ($i = 0; $i -lt 2; $i++) { [GC]::Collect(); [GC]::WaitForPendingFinalizers() }
    if (-not $ExcelPid) { return 'completed' }
    $deadline = (Get-Date).AddSeconds($ExitWaitSeconds)
    while ((Get-Process -Id $ExcelPid -ErrorAction SilentlyContinue) -and (Get-Date) -lt $deadline) {
        Start-Sleep -Milliseconds 250
    }
    if (Get-Process -Id $ExcelPid -ErrorAction SilentlyContinue) {
        Stop-Process -Id $ExcelPid -Force
        return "forced: Excel still running $ExitWaitSeconds s after Quit"
    }
    return 'completed'
}

function Start-Excel {
    $excel = New-Object -ComObject Excel.Application
    [uint32] $pidValue = 0
    [void][CpmBuildWin32]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$pidValue)
    $excel.DisplayAlerts = $false
    $excel.Visible = $true
    return @($excel, $pidValue)
}

function Read-Zip([string] $Path) {
    # Every entry in order, as name and bytes.
    $entries = New-Object System.Collections.Generic.List[object]
    $archive = [IO.Compression.ZipFile]::OpenRead($Path)
    try {
        foreach ($entry in $archive.Entries) {
            $stream = $entry.Open()
            try {
                $buffer = New-Object IO.MemoryStream
                $stream.CopyTo($buffer)
                $entries.Add([pscustomobject]@{ Name = $entry.FullName; Bytes = $buffer.ToArray() })
            } finally { $stream.Dispose() }
        }
    } finally { $archive.Dispose() }
    return $entries
}

function Get-MetadataProblems($Entries) {
    # The same kinds tools\release_provenance.py refuses (#191).
    $problems = @()
    $utf8 = New-Object Text.UTF8Encoding($false)
    foreach ($entry in $Entries) {
        $name = $entry.Name
        if ($name -like 'docMetadata/*') { $problems += "sensitivity label part: $name"; continue }
        if ($name -notmatch '\.(xml|rels)$') { continue }
        $text = $utf8.GetString($entry.Bytes)
        if ($name -eq '_rels/.rels' -and $text -match 'classificationlabels') { $problems += 'sensitivity label relationship in _rels/.rels' }
        if ($name -eq 'docProps/custom.xml' -and $text -match 'MSIP_Label_') { $problems += 'sensitivity-label properties in docProps/custom.xml' }
        if ($name -eq 'docProps/core.xml' -and $text -match '<cp:(category|contentStatus)\b[^>]*>\s*[^<\s]') { $problems += "document $($Matches[1]) in docProps/core.xml" }
        if ($name -eq 'xl/workbook.xml' -and $text -match 'absPath') { $problems += 'saved-from folder in xl/workbook.xml' }
        if ($name -match '^xl/(worksheets|chartsheets|dialogsheets|macrosheets)/[^/]+\.xml$' -and
            $text -match '<(odd|even|first)(Header|Footer)\b[^>]*>\s*[^<\s]') { $problems += "print header or footer in $name" }
    }
    return $problems
}

function Remove-Metadata([string] $Path) {
    # Rewrites the workbook without the managed-machine metadata; every other
    # entry keeps its bytes and its order. Returns what was removed.
    $utf8 = New-Object Text.UTF8Encoding($false)
    $removed = @()
    $kept = New-Object System.Collections.Generic.List[object]
    foreach ($entry in (Read-Zip $Path)) {
        $name = $entry.Name
        if ($name -like 'docMetadata/*') { $removed += $name; continue }
        $bytes = $entry.Bytes
        if ($name -match '\.(xml|rels)$') {
            $text = $utf8.GetString($bytes)
            $new = $text
            if ($name -eq '_rels/.rels') { $new = $new -replace '<Relationship [^>]*classificationlabels[^>]*/>', '' }
            if ($name -eq '[Content_Types].xml') { $new = $new -replace '<Override PartName="/docMetadata/[^"]*"[^>]*/>', '' }
            if ($name -eq 'docProps/core.xml') { $new = $new -replace '<cp:(category|contentStatus)\b[^>]*>[^<]*</cp:\1>', '' }
            if ($name -eq 'docProps/custom.xml') { $new = $new -replace '(?s)<property [^>]*name="MSIP_Label_[^"]*"[^>]*>.*?</property>', '' }
            if ($name -eq 'xl/workbook.xml') {
                $new = $new -replace '(?s)<mc:AlternateContent\b[^>]*>(?:(?!</mc:AlternateContent>).)*?absPath(?:(?!</mc:AlternateContent>).)*?</mc:AlternateContent>', ''
            }
            if ($name -match '^xl/worksheets/[^/]+\.xml$') { $new = $new -replace '(?s)<headerFooter\b[^>]*>.*?</headerFooter>', '' }
            if ($new -ne $text) { $removed += $name; $bytes = $utf8.GetBytes($new) }
        }
        $kept.Add([pscustomobject]@{ Name = $name; Bytes = $bytes })
    }
    if ($removed.Count -eq 0) { return $removed }
    Remove-Item -LiteralPath $Path
    $archive = [IO.Compression.ZipFile]::Open($Path, [IO.Compression.ZipArchiveMode]::Create)
    try {
        foreach ($entry in $kept) {
            $stream = $archive.CreateEntry($entry.Name, [IO.Compression.CompressionLevel]::Optimal).Open()
            try { $stream.Write($entry.Bytes, 0, $entry.Bytes.Length) } finally { $stream.Dispose() }
        }
    } finally { $archive.Dispose() }
    return $removed
}

function Get-SheetNames($Book) {
    # Names in workbook order. Each sheet object is released at once: a COM
    # reference left in a variable keeps Excel alive after Quit.
    $names = @()
    $sheets = $Book.Worksheets
    try {
        for ($i = 1; $i -le $sheets.Count; $i++) {
            $sheet = $sheets.Item($i)
            try { $names += [string]$sheet.Name } finally { Release-Com $sheet; $sheet = $null }
        }
    } finally { Release-Com $sheets }
    return $names
}

function Get-Sha256([string] $Path) { return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLower() }

# -- Preconditions ------------------------------------------------------------
$sha = (& git -C $repo rev-parse HEAD | Out-String).Trim()
if ($LASTEXITCODE -ne 0 -or $sha -notmatch '^[0-9a-f]{40}$') { throw 'Run from a Git checkout of the repository.' }
$dirty = (& git -C $repo status --porcelain | Out-String).Trim()
if ($dirty) { throw "The checkout has local changes; build only from a clean commit:`n$dirty" }
$version = (Get-Content -Raw -LiteralPath (Join-Path $repo 'VERSION')).Trim()
Stop-WindowlessExcel
$stamp = (Get-Date).ToUniversalTime()
$runDir = Join-Path $OutDir ('demo-' + $sha.Substring(0, 7) + '-' + $stamp.ToString('yyyyMMddTHHmmssZ'))
[void](New-Item -ItemType Directory -Path $runDir -Force)
$workbookPath = Join-Path $runDir $WorkbookName
$packagePath = Join-Path $runDir ("VBA-Performance-Manager-v$version.zip")
Write-Host "Commit:  $sha"
Write-Host "Version: $version"
Write-Host "Output:  $runDir"

$record = [ordered]@{
    tool = 'Build-DemoWorkbook.ps1'; commit_sha = $sha; version = $version
    started_utc = $stamp.ToString('o'); state = 'not started'; excel = $null
    imported = @(); compiled = $false; sheets = @(); actions = @(); failure_count = $null
    metadata_removed = @(); metadata_problems = @(); reopened_sheets = @()
    workbook = $null; package = $null; cleanup = $null; error = $null; reasons = @(); passed = $false
}

$vbomExisted = $false; $vbomBefore = $null; $vbomTouched = $false
$excel = $null; $excelPid = 0; $workbook = $null; $project = $null; $control = $null
$sheet = $null; $table = $null; $tables = $null; $logSheetObject = $null; $homeSheet = $null
try {
    if ($AllowVbomAccess) {
        if (-not (Test-Path $VbomKey)) { [void](New-Item -Path $VbomKey -Force) }
        $current = Get-ItemProperty -Path $VbomKey -Name AccessVBOM -ErrorAction SilentlyContinue
        if ($null -ne $current) { $vbomExisted = $true; $vbomBefore = $current.AccessVBOM }
        $vbomTouched = $true
        Set-ItemProperty -Path $VbomKey -Name AccessVBOM -Value 1 -Type DWord
    }

    # -- Build -------------------------------------------------------------
    $started = Start-Excel; $excel = $started[0]; $excelPid = $started[1]
    $record.excel = [ordered]@{ version = $excel.Version; build = $excel.Build }
    $workbook = $excel.Workbooks.Add()
    try { $project = $workbook.VBProject; [void]$project.VBComponents.Count }
    catch { throw 'Excel refused access to the VBA project. Run the .cmd launcher, which passes -AllowVbomAccess.' }

    $record.state = 'importing'
    foreach ($relative in $Components) {
        $file = Join-Path $repo $relative
        [void]$project.VBComponents.Import($file)
        $record.imported += [ordered]@{ component = $relative; sha256 = (Get-Sha256 $file) }
    }

    $record.state = 'compiling'
    $control = $excel.VBE.CommandBars.FindControl([Reflection.Missing]::Value, $CompileControlId)
    if ($null -eq $control) { throw 'Could not find the Compile VBAProject command.' }
    if ($control.Enabled) { $control.Execute() }
    if ($control.Enabled) { throw 'The VBA project did not compile.' }
    $record.compiled = $true

    $record.state = 'building sheets'
    $prefix = "'$($workbook.Name)'!"
    [void]$excel.Run($prefix + $BuildEntry)
    foreach ($name in @(Get-SheetNames $workbook)) {
        if ($Sheets -contains $name) { continue }
        $sheet = $workbook.Worksheets.Item($name)
        try { $sheet.Delete() } finally { Release-Com $sheet; $sheet = $null }
    }
    $record.sheets = @(Get-SheetNames $workbook)
    if (($record.sheets -join ',') -ne ($Sheets -join ',')) {
        throw "Unexpected sheets after the build: $($record.sheets -join ', ')"
    }

    # -- Smoke test --------------------------------------------------------
    $record.state = 'smoke test'
    [void]$excel.Run($prefix + 'cPM_Demo_SetUnattended', $true)
    $logSheetObject = $workbook.Worksheets.Item($LogSheet)
    $tables = $logSheetObject.ListObjects
    $table = $tables.Item($LogTable)
    foreach ($action in $SmokeActions) {
        $before = [int]$table.ListRows.Count
        $failuresBefore = [int]$excel.Run($prefix + 'cPM_Demo_FailureCount')
        $clock = [Diagnostics.Stopwatch]::StartNew()
        [void]$excel.Run($prefix + $action)
        $clock.Stop()
        $added = [int]$table.ListRows.Count - $before
        $failed = [int]$excel.Run($prefix + 'cPM_Demo_FailureCount') - $failuresBefore
        $record.actions += [ordered]@{ action = $action; rows_added = $added; failed = ($failed -gt 0)
                                       seconds = [math]::Round($clock.Elapsed.TotalSeconds, 2) }
        Write-Host ("  {0,-26} rows {1,3}  {2}" -f $action, $added, $(if ($failed -gt 0 -or $added -lt 1) { 'FAILED' } else { 'ok' }))
    }
    $record.failure_count = [int]$excel.Run($prefix + 'cPM_Demo_FailureCount')
    [void]$excel.Run($prefix + $ClearAction)
    if ([int]$table.ListRows.Count -ne 0) { throw 'The results log is not empty after cPM_ClearResultsLog.' }
    Release-Com $table; Release-Com $tables; Release-Com $logSheetObject
    $table = $null; $tables = $null; $logSheetObject = $null
    [void]$excel.Run($prefix + 'cPM_Demo_SetUnattended', $false)
    $homeSheet = $workbook.Worksheets.Item('Home')
    try { $homeSheet.Activate() } finally { Release-Com $homeSheet; $homeSheet = $null }

    $record.state = 'saving'
    $workbook.SaveAs($workbookPath, 52)
    Release-Com $control; Release-Com $project; $control = $null; $project = $null
    $workbook.Close($false); Release-Com $workbook; $workbook = $null
    $record.cleanup = Close-Excel $excel $excelPid; $excel = $null; $excelPid = 0

    # -- Metadata ------------------------------------------------------------
    $record.state = 'removing metadata'
    $record.metadata_removed = @(Remove-Metadata $workbookPath)
    $record.metadata_problems = @(Get-MetadataProblems (Read-Zip $workbookPath))

    # -- Reopen ----------------------------------------------------------------
    $record.state = 'reopening'
    Stop-WindowlessExcel
    $started = Start-Excel; $excel = $started[0]; $excelPid = $started[1]
    $excel.AutomationSecurity = 3
    $workbook = $excel.Workbooks.Open($workbookPath, 0, $true)
    $record.reopened_sheets = @(Get-SheetNames $workbook)
    $workbook.Close($false); Release-Com $workbook; $workbook = $null
    $reopenCleanup = Close-Excel $excel $excelPid; $excel = $null; $excelPid = 0
    if ($reopenCleanup -ne 'completed') { $record.cleanup = $reopenCleanup }

    # -- Package -------------------------------------------------------------
    $record.state = 'packaging'
    $archive = [IO.Compression.ZipFile]::Open($packagePath, [IO.Compression.ZipArchiveMode]::Create)
    try {
        foreach ($member in @($WorkbookName, 'LICENSE', 'NOTICE')) {
            $source = $(if ($member -eq $WorkbookName) { $workbookPath } else { Join-Path $repo $member })
            [void][IO.Compression.ZipFileExtensions]::CreateEntryFromFile($archive, $source, $member,
                                                                         [IO.Compression.CompressionLevel]::Optimal)
        }
    } finally { $archive.Dispose() }
    $record.workbook = [ordered]@{ file = $WorkbookName; sha256 = (Get-Sha256 $workbookPath)
                                   size = (Get-Item -LiteralPath $workbookPath).Length }
    $record.package = [ordered]@{ file = (Split-Path -Leaf $packagePath); sha256 = (Get-Sha256 $packagePath)
                                  members = [ordered]@{ $WorkbookName = $record.workbook.sha256
                                                        LICENSE = (Get-Sha256 (Join-Path $repo 'LICENSE'))
                                                        NOTICE = (Get-Sha256 (Join-Path $repo 'NOTICE')) } }
    $record.state = 'completed'
} catch {
    $record.error = $_.Exception.Message
    Write-Host "FAIL: $($record.error)"
} finally {
    Release-Com $homeSheet; Release-Com $sheet; Release-Com $table; Release-Com $tables; Release-Com $logSheetObject
    Release-Com $control; Release-Com $project
    if ($null -ne $workbook) { Release-Com $workbook }
    if ($null -ne $excel -or $excelPid) { $record.cleanup = Close-Excel $excel $excelPid }
    if ($vbomTouched) {
        try {
            if ($vbomExisted) { Set-ItemProperty -Path $VbomKey -Name AccessVBOM -Value $vbomBefore -Type DWord }
            else { Remove-ItemProperty -Path $VbomKey -Name AccessVBOM -ErrorAction Stop }
        } catch { $record.cleanup = 'failed to restore AccessVBOM: ' + $_.Exception.Message }
    }
}

# Every reason the build is not publishable, so a FAIL always says why.
$reasons = @()
if ($record.state -ne 'completed') { $reasons += "stopped while $($record.state): $($record.error)" }
if (-not $record.compiled) { $reasons += 'the VBA project was not compiled' }
if ($record.actions.Count -ne $SmokeActions.Count) { $reasons += "$($record.actions.Count) of $($SmokeActions.Count) demo actions ran" }
foreach ($action in @($record.actions | Where-Object { $_.failed -or $_.rows_added -lt 1 })) {
    $reasons += "demo action $($action.action) failed or added no log row"
}
if ($record.failure_count -ne 0) { $reasons += "the demo reported $($record.failure_count) failure(s)" }
foreach ($problem in $record.metadata_problems) { $reasons += "metadata left: $problem" }
if (($record.reopened_sheets -join ',') -ne ($Sheets -join ',')) {
    $reasons += "reopened sheets are '$($record.reopened_sheets -join ', ')', expected '$($Sheets -join ', ')'"
}
if ($record.cleanup -ne 'completed') { $reasons += "Excel cleanup: $($record.cleanup)" }
$record.reasons = $reasons
$record.passed = ($reasons.Count -eq 0)
$record.finished_utc = (Get-Date).ToUniversalTime().ToString('o')
$recordPath = Join-Path $runDir 'demo-build.json'
$record | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $recordPath -Encoding UTF8
Write-Host ''
if ($record.metadata_removed.Count) { Write-Host "Metadata removed from: $($record.metadata_removed -join ', ')" }
if ($record.workbook) { Write-Host "Workbook SHA-256: $($record.workbook.sha256)" }
if ($record.package) { Write-Host "Package:          $($record.package.file)  $($record.package.sha256)" }
Write-Host "Build record:     $recordPath"
if ($record.passed) { Write-Host 'RESULT=PASS'; exit 0 }
foreach ($reason in $reasons) { Write-Host "  - $reason" }
Write-Host 'RESULT=FAIL; do not publish this workbook.'
exit 1
