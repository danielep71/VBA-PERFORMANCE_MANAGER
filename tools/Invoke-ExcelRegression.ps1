<#
.SYNOPSIS
    Runs the cPerformanceManager regression suite in desktop Excel from an exact,
    clean Git checkout, and writes a JSON evidence record (#11, local runner).

.DESCRIPTION
    For each run the script:
      1. checks that the checkout is at -ExpectedSha and has no local changes;
      2. starts its own Excel instance and adds a new, empty workbook;
      3. imports the four components in the documented order;
      4. compiles the project (Debug > Compile VBAProject);
      5. runs Run_cPerformanceManager_RegressionSuite;
      6. reads every case from the REGRESSION_cPM worksheet;
      7. closes the workbook without saving and quits that Excel instance.

    With -RedFile it then repeats the run with that file in place of
    -RedComponent (default src\modules\M_cPM_TIMEWASTERS.bas), and expects
    exactly case -RedExpectedCase to fail. This is how a change shows that its
    regression detects the defect it fixes. A red run is never certification
    evidence.

    The evidence record holds "result" for the source run in the v1.4.0 frozen
    execution-result schema (commit_sha, run_state, excel_version_build,
    office_bitness, cases, assertions, failures, cleanup_state, timestamp_utc),
    and "details" with every run, the imported files' SHA-256 and the red check.

    Limits of this local runner: it has no timeout, and a VBA compile error or
    an unhandled runtime error opens a modal Excel dialog. Excel stays visible
    so the maintainer can dismiss it; the run is then recorded as failed.

    Importing VBA through automation needs Excel's "Trust access to the VBA
    project object model" (AccessVBOM). With -AllowVbomAccess the script turns
    it on for the current Windows user only, for the duration of the run, and
    puts back the previous value afterwards, even on failure. Without the
    switch it changes no setting.

    The script never saves a workbook, never pushes, and changes nothing in
    the repository.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File tools\Invoke-ExcelRegression.ps1 `
        -ExpectedSha <full commit SHA> -AllowVbomAccess -OutDir ..\excel-evidence

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File tools\Invoke-ExcelRegression.ps1 `
        -ExpectedSha <full commit SHA> -AllowVbomAccess -OutDir ..\excel-evidence `
        -RedFile ..\red\M_cPM_TIMEWASTERS.bas -RedExpectedCase 81
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)] [string] $ExpectedSha,
    [string] $Repo,
    [string] $OutDir,
    [string] $RedFile,
    [string] $RedComponent = 'src\modules\M_cPM_TIMEWASTERS.bas',
    [int] $RedExpectedCase = 0,
    [switch] $AllowVbomAccess
)

Set-StrictMode -Version 2
$ErrorActionPreference = 'Stop'

# The documented import order (README, Regression testing).
$Components = @(
    'src\modules\M_cPM_TIMEWASTERS.bas',
    'src\classes\cPerformanceManager.cls',
    'demo\M_DEMO_BUILDER.bas',
    'test\M_cPM_Test.bas'
)
$EntryPoint = 'Run_cPerformanceManager_RegressionSuite'
$LogSheet = 'REGRESSION_cPM'
$FirstSummaryRow = 6            # the suite writes one row per case from row 6
$CompileControlId = 578         # Debug > Compile VBAProject
$VbomKey = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security'

if (-not $Repo) { $Repo = Split-Path -Parent $PSScriptRoot }
$Repo = (Resolve-Path $Repo).Path
if (-not $OutDir) { throw '-OutDir is required: a folder outside the repository for the evidence record.' }
# Resolve against PowerShell's current location, which .NET's GetFullPath ignores.
$OutDir = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutDir)
if (($OutDir.TrimEnd('\') + '\').StartsWith($Repo.TrimEnd('\') + '\', [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "-OutDir must be outside the repository ($Repo), or the evidence would dirty the checkout."
}
if ($RedFile) {
    $RedFile = (Resolve-Path $RedFile).Path
    if ($Components -notcontains $RedComponent) { throw "-RedComponent must be one of: $($Components -join ', ')" }
    if ($RedExpectedCase -lt 1) { throw '-RedFile needs -RedExpectedCase, the one case expected to fail.' }
}

Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class CpmWin32 {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
'@

function Invoke-Git([string[]] $Arguments) {
    # No 2>&1: in Windows PowerShell 5.1 a redirected stderr line would stop
    # the script under ErrorActionPreference Stop. The exit code is checked.
    $output = & git -C $Repo @Arguments
    if ($LASTEXITCODE -ne 0) { throw "git $($Arguments -join ' ') failed with exit code $LASTEXITCODE" }
    return ($output | Out-String).Trim()
}

function Get-Bitness([string] $ExePath) {
    # Read the machine field of the executable's PE header.
    $stream = [System.IO.File]::OpenRead($ExePath)
    try {
        $reader = New-Object System.IO.BinaryReader($stream)
        $stream.Position = 0x3C
        $stream.Position = $reader.ReadInt32() + 4
        switch ($reader.ReadUInt16()) {
            0x8664 { return '64-bit' }
            0x014C { return '32-bit' }
            0xAA64 { return 'ARM64' }
            default { return 'unknown' }
        }
    } finally { $stream.Dispose() }
}

function Get-OfficeBuild {
    $key = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
    try { return (Get-ItemProperty -Path $key -Name VersionToReport).VersionToReport }
    catch { return $null }
}

function Release-Com($object) {
    if ($null -ne $object) {
        try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($object) } catch { }
    }
}

function Invoke-Suite([string] $Label, [string] $ReplacedComponent, [string] $ReplacementFile) {
    $result = [ordered]@{
        run = $Label; started = (Get-Date).ToUniversalTime().ToString('o')
        state = 'not started'; imported = @(); compiled = $false
        excel = $null; cases = 0; assertions = 0; failures = 0
        failed_cases = @(); cleanup = $null; error = $null
    }
    $excel = $null; $workbook = $null; $excelPid = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $true
        $excel.DisplayAlerts = $false
        [uint32] $processId = 0
        [void][CpmWin32]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref] $processId)
        $excelPid = $processId
        $exePath = (Get-Process -Id $excelPid).Path
        $result.excel = [ordered]@{
            version = $excel.Version; build = $excel.Build
            version_to_report = Get-OfficeBuild; bitness = Get-Bitness $exePath; path = $exePath
        }

        $workbook = $excel.Workbooks.Add()
        try { $project = $workbook.VBProject; [void]$project.VBComponents.Count }
        catch {
            throw ("Excel refused access to the VBA project. Run again with -AllowVbomAccess, " +
                   "or check whether a policy disables 'Trust access to the VBA project object model'.")
        }

        $result.state = 'importing'
        foreach ($relative in $Components) {
            $file = Join-Path $Repo $relative
            if ($ReplacementFile -and $relative -eq $ReplacedComponent) { $file = $ReplacementFile }
            [void]$project.VBComponents.Import($file)
            $result.imported += [ordered]@{
                component = $relative; file = $file
                sha256 = (Get-FileHash -Path $file -Algorithm SHA256).Hash.ToLower()
            }
        }

        $result.state = 'compiling'
        $control = $excel.VBE.CommandBars.FindControl([System.Reflection.Missing]::Value, $CompileControlId)
        if ($null -eq $control) { throw 'Could not find the Compile VBAProject command.' }
        if ($control.Enabled) { $control.Execute() }
        # The command stays enabled while the project does not compile.
        if ($control.Enabled) { throw 'The VBA project did not compile.' }
        $result.compiled = $true
        Release-Com $control

        $result.state = 'running'
        [void]$excel.Run("'$($workbook.Name)'!$EntryPoint")

        $result.state = 'reading results'
        $sheet = $workbook.Worksheets.Item($LogSheet)
        $row = $FirstSummaryRow
        while ($null -ne $sheet.Cells.Item($row, 4).Value2) {
            $case = [int]$sheet.Cells.Item($row, 4).Value2
            $assertions = [int]$sheet.Cells.Item($row, 6).Value2
            $failures = [int]$sheet.Cells.Item($row, 7).Value2
            $outcome = [string]$sheet.Cells.Item($row, 8).Value2
            $result.cases += 1
            $result.assertions += $assertions
            $result.failures += $failures
            if ($failures -gt 0 -or $outcome -ne 'PASS') {
                $result.failed_cases += [ordered]@{
                    case = $case; name = [string]$sheet.Cells.Item($row, 5).Value2
                    failures = $failures; result = $outcome
                }
            }
            $row += 1
        }
        Release-Com $sheet
        $result.state = 'completed'
    } catch {
        $result.error = $_.Exception.Message
    } finally {
        if ($null -ne $workbook) { try { $workbook.Close($false) } catch { } }
        if ($null -ne $excel) { try { $excel.Quit() } catch { } }
        Release-Com $workbook; Release-Com $excel
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        $result.cleanup = 'completed'
        if ($excelPid) {
            Start-Sleep -Seconds 3
            $left = Get-Process -Id $excelPid -ErrorAction SilentlyContinue
            if ($left) { Stop-Process -Id $excelPid -Force; $result.cleanup = 'forced: Excel stopped after Quit' }
        }
        $result.finished = (Get-Date).ToUniversalTime().ToString('o')
    }
    return $result
}

# --------------------------------------------------------------------------- #
# Preconditions
# --------------------------------------------------------------------------- #

$sha = Invoke-Git @('rev-parse', 'HEAD')
$branch = Invoke-Git @('rev-parse', '--abbrev-ref', 'HEAD')
$dirty = Invoke-Git @('status', '--porcelain')
if (-not $sha.StartsWith($ExpectedSha.ToLower())) { throw "The checkout is at $sha, not $ExpectedSha." }
if ($dirty) { throw "The checkout has local changes:`n$dirty" }
if (Get-Process -Name EXCEL -ErrorAction SilentlyContinue) {
    throw 'Excel is already running. Save your work, close every Excel window, then run again.'
}
$testText = Get-Content -Raw -Path (Join-Path $Repo 'test\M_cPM_Test.bas')
if ($testText -notmatch 'TotalSteps\s+As Long\s*=\s*(\d+)') { throw 'TotalSteps not found in the test module.' }
$expectedCases = [int]$Matches[1]

# --------------------------------------------------------------------------- #
# Runs, with AccessVBOM restored on every path
# --------------------------------------------------------------------------- #

$vbomBefore = $null; $vbomExisted = $false
if ($AllowVbomAccess) {
    if (-not (Test-Path $VbomKey)) { [void](New-Item -Path $VbomKey -Force) }
    $current = Get-ItemProperty -Path $VbomKey -Name AccessVBOM -ErrorAction SilentlyContinue
    if ($null -ne $current) { $vbomExisted = $true; $vbomBefore = $current.AccessVBOM }
    Set-ItemProperty -Path $VbomKey -Name AccessVBOM -Value 1 -Type DWord
}
$runs = @()
try {
    Write-Host "Run 1 of $(if ($RedFile) { 2 } else { 1 }): source at $sha"
    $runs += Invoke-Suite -Label 'source' -ReplacedComponent '' -ReplacementFile ''
    if ($RedFile) {
        Write-Host "Run 2 of 2: red check with $RedFile in place of $RedComponent"
        $runs += Invoke-Suite -Label 'red check' -ReplacedComponent $RedComponent -ReplacementFile $RedFile
    }
} finally {
    if ($AllowVbomAccess) {
        if ($vbomExisted) { Set-ItemProperty -Path $VbomKey -Name AccessVBOM -Value $vbomBefore -Type DWord }
        else { Remove-ItemProperty -Path $VbomKey -Name AccessVBOM -ErrorAction SilentlyContinue }
    }
}

# --------------------------------------------------------------------------- #
# Verdicts and evidence
# --------------------------------------------------------------------------- #

$source = $runs[0]
$sourcePass = ($source.state -eq 'completed' -and $source.cases -eq $expectedCases -and
               $source.failures -eq 0 -and $source.cleanup -eq 'completed')
$verdicts = [ordered]@{ source = $(if ($sourcePass) { 'PASS' } else { 'FAIL' }) }
if ($RedFile) {
    $red = $runs[1]
    $redCases = @($red.failed_cases | ForEach-Object { $_.case })
    $redAsExpected = ($red.state -eq 'completed' -and $red.cases -eq $expectedCases -and
                      $redCases.Count -eq 1 -and $redCases[0] -eq $RedExpectedCase)
    $verdicts.red_check = $(if ($redAsExpected) { "AS EXPECTED: only case $RedExpectedCase fails" } else { 'NOT AS EXPECTED' })
}
# "result" is the v1.4.0 frozen execution-result schema, for the source run only.
$excelBuild = 'unknown'
if ($source.excel) {
    $excelBuild = "Microsoft Excel $($source.excel.version) build $($source.excel.build)"
    if ($source.excel.version_to_report) { $excelBuild += " ($($source.excel.version_to_report))" }
}
$record = [ordered]@{
    result = [ordered]@{
        commit_sha = $sha
        run_state = $source.state
        excel_version_build = $excelBuild
        office_bitness = $(if ($source.excel) { $source.excel.bitness } else { 'unknown' })
        cases = $source.cases
        assertions = $source.assertions
        failures = $source.failures
        cleanup_state = $source.cleanup
        timestamp_utc = $source.finished
    }
    details = [ordered]@{
        tool = 'tools/Invoke-ExcelRegression.ps1'; branch = $branch; clean = $true
        expected_cases = $expectedCases; vbom_changed = [bool]$AllowVbomAccess
        verdicts = $verdicts; runs = $runs
    }
}
if (-not (Test-Path $OutDir)) { [void](New-Item -ItemType Directory -Path $OutDir -Force) }
$stamp = (Get-Date).ToUniversalTime().ToString('yyyyMMddTHHmmssZ')
$path = Join-Path $OutDir "cpm-regression-$($sha.Substring(0, 7))-$stamp.json"
$record | ConvertTo-Json -Depth 6 | Set-Content -Path $path -Encoding UTF8

foreach ($run in $runs) {
    Write-Host ''
    Write-Host "[$($run.run)] state: $($run.state)"
    if ($run.excel) { Write-Host "  Excel $($run.excel.version) build $($run.excel.build) ($($run.excel.version_to_report)), $($run.excel.bitness)" }
    Write-Host "  cases $($run.cases) of $expectedCases, assertions $($run.assertions), failures $($run.failures)"
    foreach ($failed in $run.failed_cases) { Write-Host "  FAIL case $($failed.case): $($failed.name)" }
    if ($run.error) { Write-Host "  error: $($run.error)" }
}
Write-Host ''
foreach ($key in $verdicts.Keys) { Write-Host "$key : $($verdicts[$key])" }
Write-Host "Evidence written to $path"

$ok = $sourcePass -and ((-not $RedFile) -or $redAsExpected)
if ($ok) { exit 0 } else { exit 1 }
