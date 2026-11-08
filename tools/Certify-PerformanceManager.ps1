<#
.SYNOPSIS
Run the checked-out Performance Manager candidate with a double click.
.DESCRIPTION
Detects the clean checkout's SHA and creates a unique evidence directory next
to it. Invokes the existing Excel runner in a separate PowerShell process,
records session.log and checks the JSON verdict before reporting PASS.
Requires Git and desktop Excel on Windows. No Python or administrator needed.
The runner temporarily enables VBA project access and restores its previous
value. No macro policy is changed. Existing Excel processes must be closed.
This remains a supervised local run: modal Excel errors have no timeout.
#>
[CmdletBinding()]
param([switch] $TimerRollbackRedCheck, [switch] $MeasurementExamples)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$bundle = $null
$transcribing = $false
$code = 1
try {
    if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
        throw 'Git is required. Use a Git checkout, not a downloaded ZIP.'
    }
    $sha = (& git -C $repo rev-parse HEAD | Out-String).Trim()
    if ($LASTEXITCODE -ne 0 -or $sha -notmatch '^[0-9a-f]{40}$') {
        throw 'Cannot identify the checked-out commit.'
    }
    $dirty = (& git -C $repo status --porcelain | Out-String).Trim()
    if ($LASTEXITCODE -ne 0 -or $dirty) {
        throw 'The checkout has local changes. Commit or move them before certification.'
    }
    if (Get-Process -Name EXCEL -ErrorAction SilentlyContinue) {
        throw 'Save your work and close every Excel window, then double-click again.'
    }
    $name = 'cpm-evidence-' + $sha.Substring(0, 7) + '-' + [Guid]::NewGuid().ToString('N')
    $bundle = Join-Path (Split-Path -Parent $repo) $name
    [void](New-Item -ItemType Directory -Path $bundle)
    Start-Transcript -Path (Join-Path $bundle 'session.log') | Out-Null
    $transcribing = $true
    Write-Host "Candidate: $sha"
    Write-Host 'Excel will open, import, compile and run the full suite automatically.'
    Write-Host 'VBA project access is enabled temporarily and restored by the runner.'
    Write-Host 'If Excel displays an error dialog, capture it before dismissing it.'
    $runner = Join-Path $PSScriptRoot 'Invoke-ExcelRegression.ps1'
    $redArgs = @()
    if ($MeasurementExamples) { $redArgs += '-MeasurementExamples' }
    if ($TimerRollbackRedCheck) {
        $classPath = Join-Path $repo 'src\classes\cPerformanceManager.cls'
        $sourceText = [System.IO.File]::ReadAllText($classPath)
        $pattern = '(?m)^(\s*)Start_ReleaseAcquiredPeriod PeriodHeldBefore\r?$'
        if ([regex]::Matches($sourceText, $pattern).Count -ne 1) {
            throw 'Expected exactly one rollback call; cannot construct the negative control.'
        }
        $redText = [regex]::Replace($sourceText, $pattern, '$1'' Negative control: rollback call disabled')
        $redPath = Join-Path $bundle 'cPerformanceManager-red.cls'
        [System.IO.File]::WriteAllText($redPath, $redText, [System.Text.Encoding]::ASCII)
        $redArgs += @('-RedFile', $redPath, '-RedComponent', 'src\classes\cPerformanceManager.cls', '-RedExpectedCase', '86')
        Write-Host 'Two runs: source must pass; negative control must fail only case 86.'
    }
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $runner `
        -ExpectedSha $sha -Repo $repo -OutDir $bundle -AllowVbomAccess @redArgs
    $runnerCode = $LASTEXITCODE
    $records = @(Get-ChildItem -LiteralPath $bundle -Filter 'cpm-regression-*.json')
    if ($runnerCode -ne 0 -or $records.Count -ne 1) {
        throw 'The Excel run did not pass or its JSON record is missing. Keep the evidence folder.'
    }
    $record = Get-Content -LiteralPath $records[0].FullName -Raw | ConvertFrom-Json
    if ($record.result.commit_sha -ne $sha -or
        $record.result.run_state -ne 'completed' -or
        $record.result.office_bitness -ne '64-bit' -or
        $record.result.cases -ne $record.details.expected_cases -or
        $record.result.cases -le 0 -or $record.result.failures -ne 0 -or
        $record.result.cleanup_state -ne 'completed' -or
        $record.details.verdicts.source -ne 'PASS') {
        throw 'The evidence does not certify a completed, clean 64-bit Excel run.'
    }
    if ($MeasurementExamples -and $record.details.runs[0].examples.passed -ne $true) {
        throw 'Measurement examples did not pass; keep the transcript in the JSON.'
    }
    if ($TimerRollbackRedCheck) {
        $runs = @($record.details.runs)
        if ($runs.Count -ne 2 -or $runs[1].cleanup -ne 'completed' -or
            $runs[1].compiled -ne $true -or
            $record.details.verdicts.red_check -ne 'AS EXPECTED: only case 86 fails') {
            throw 'The negative control did not finish with the expected failure and clean Excel exit.'
        }
        Write-Host 'NEGATIVE CONTROL=PASS: only case 86 fails without the fix; cleanup completed.'
    }
    Write-Host "RESULT=PASS: $($record.result.cases) cases, $($record.result.assertions) assertions, zero failures."
    Write-Host 'Send the JSON file in the evidence folder for review.'
    $code = 0
} catch {
    Write-Host ('RESULT=FAIL: ' + $_.Exception.Message)
} finally {
    if ($transcribing) { Stop-Transcript | Out-Null }
    if ($bundle) {
        Write-Host "Evidence folder: $bundle"
        try { Invoke-Item -LiteralPath $bundle } catch { Write-Host 'Open the evidence folder manually.' }
    }
}
exit $code
