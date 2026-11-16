<#
.SYNOPSIS
    Records evidence that the workbook sync refuses a protected VBA project and
    a component type conflict, restores AccessVBOM and leaves no Excel process
    behind (#138).

.DESCRIPTION
    Supervised Windows desktop Excel tool, started by Verify-SyncRefusals.cmd.
    It works in a folder next to the checkout (default
    ..\cpm-sync-refusals) and never changes the checkout.

      1. conflict.xlsm is created automatically: a new workbook whose class
         module is named M_cPM_TIMEWASTERS, which the sync expects as a
         standard module.
      2. protected.xlsm is created automatically with one standard module.
         Excel cannot lock a VBA project through automation, so the script
         opens it with the VBA editor and asks for one manual step: lock the
         project with a password, save, close the workbook. Later runs reuse
         the locked file.
      3. tools\Manage-PerformanceManagerVBA.ps1 -Mode Sync -AllowVbomAccess
         runs on each workbook. Each run must fail with the expected refusal,
         leave the input unchanged, report cleanup 'completed', restore
         AccessVBOM to its prior state and leave no Excel process.

    The evidence is written to sync-refusals-<sha>-<UTC time>.json in the
    work folder. Excel windows without a title (invisible automation
    instances) are closed first; a visible Excel window stops the script.
#>
[CmdletBinding()]
param(
    [string] $WorkDir
)
Set-StrictMode -Version 2
$ErrorActionPreference = 'Stop'

$repo = Split-Path -Parent $PSScriptRoot
$tool = Join-Path $PSScriptRoot 'Manage-PerformanceManagerVBA.ps1'
$vbomKey = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security'
if (-not $WorkDir) { $WorkDir = Join-Path (Split-Path -Parent $repo) 'cpm-sync-refusals' }
$WorkDir = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($WorkDir)
if ($WorkDir.TrimEnd('\').StartsWith($repo.TrimEnd('\') + '\', [StringComparison]::OrdinalIgnoreCase)) {
    throw '-WorkDir must be outside the repository.'
}

Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class CpmRefusalWin32 {
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
'@

function Release-Com($Object) {
    if ($null -ne $Object -and [Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)
    }
}

function Get-Vbom {
    $item = Get-ItemProperty -Path $vbomKey -Name AccessVBOM -ErrorAction SilentlyContinue
    if ($null -eq $item) { return 'absent' }
    return [string]$item.AccessVBOM
}

function Restore-Vbom([string] $State) {
    if ($State -eq 'absent') { Remove-ItemProperty -Path $vbomKey -Name AccessVBOM -ErrorAction SilentlyContinue }
    else { Set-ItemProperty -Path $vbomKey -Name AccessVBOM -Value ([int]$State) -Type DWord }
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

function Invoke-WithExcel([scriptblock] $Body, [switch] $Visible) {
    # Runs $Body with a private Excel instance and temporary AccessVBOM, then
    # quits that instance and restores AccessVBOM on every path.
    $vbom = Get-Vbom
    $excel = $null; $excelPid = 0
    try {
        if (-not (Test-Path $vbomKey)) { [void](New-Item -Path $vbomKey -Force) }
        Set-ItemProperty -Path $vbomKey -Name AccessVBOM -Value 1 -Type DWord
        $excel = New-Object -ComObject Excel.Application
        [uint32] $pidValue = 0
        [void][CpmRefusalWin32]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$pidValue)
        $excelPid = $pidValue
        $excel.AutomationSecurity = 3
        # The visible instance keeps its prompts, so closing an unsaved workbook asks.
        $excel.DisplayAlerts = [bool]$Visible
        $excel.Visible = [bool]$Visible
        & $Body $excel
    } finally {
        if ($null -ne $excel) {
            try { foreach ($book in @($excel.Workbooks)) { $book.Close($false); Release-Com $book } } catch { }
            try { $excel.Quit() } catch { }
            Release-Com $excel
        }
        $excel = $null
        for ($i = 0; $i -lt 2; $i++) { [GC]::Collect(); [GC]::WaitForPendingFinalizers() }
        if ($excelPid) {
            $deadline = (Get-Date).AddSeconds(15)
            while ((Get-Process -Id $excelPid -ErrorAction SilentlyContinue) -and (Get-Date) -lt $deadline) {
                Start-Sleep -Milliseconds 250
            }
            Stop-Process -Id $excelPid -Force -ErrorAction SilentlyContinue
        }
        Restore-Vbom $vbom
    }
}

function New-TestWorkbook([string] $Path, [int] $ComponentType, [string] $ComponentName) {
    Invoke-WithExcel {
        param($excel)
        $book = $excel.Workbooks.Add()
        $part = $book.VBProject.VBComponents.Add($ComponentType)
        $part.Name = $ComponentName
        if ($part.CodeModule.CountOfLines -eq 0) { $part.CodeModule.AddFromString('Option Explicit') }
        Release-Com $part
        $book.SaveAs($Path, 52)
        $book.Close($false)
        Release-Com $book
    }
}

function Get-ProjectProtection([string] $Path) {
    $script:protection = $null
    Invoke-WithExcel {
        param($excel)
        $book = $excel.Workbooks.Open($Path, 0, $true)
        $script:protection = [int]$book.VBProject.Protection
        $book.Close($false)
        Release-Com $book
    }
    return $script:protection
}

# -- Preconditions ------------------------------------------------------------
$sha = (& git -C $repo rev-parse HEAD | Out-String).Trim()
if ($LASTEXITCODE -ne 0 -or $sha -notmatch '^[0-9a-f]{40}$') { throw 'Run from a Git checkout of the repository.' }
$dirty = (& git -C $repo status --porcelain | Out-String).Trim()
if ($dirty) { throw 'The checkout is not clean; the sync would refuse it. Commit or stash your changes.' }
Stop-WindowlessExcel
if (-not (Test-Path $WorkDir)) { [void](New-Item -ItemType Directory -Path $WorkDir) }
Write-Host "Commit:      $sha"
Write-Host "Work folder: $WorkDir"

# -- Test workbooks ---------------------------------------------------------
$conflict = Join-Path $WorkDir 'conflict.xlsm'
$protected = Join-Path $WorkDir 'protected.xlsm'
if (-not (Test-Path $conflict)) {
    Write-Host 'Creating conflict.xlsm (class module named M_cPM_TIMEWASTERS)...'
    New-TestWorkbook $conflict 2 'M_cPM_TIMEWASTERS'
}
if (-not (Test-Path $protected)) {
    Write-Host 'Creating protected.xlsm (one standard module)...'
    New-TestWorkbook $protected 1 'Module1'
}
if ((Get-ProjectProtection $protected) -eq 0) {
    Write-Host ''
    Write-Host 'ONE MANUAL STEP: Excel cannot lock a VBA project through automation.'
    Write-Host 'Excel will open protected.xlsm with the VBA editor. In the editor:'
    Write-Host '  Tools > VBAProject Properties > Protection'
    Write-Host '  tick "Lock project for viewing", type any password twice, OK.'
    Write-Host 'Then save (Ctrl+S) and close the workbook window.'
    Write-Host ''
    Invoke-WithExcel {
        param($excel)
        $book = $excel.Workbooks.Open($protected)
        # PowerShell's COM adapter has no type information for the VBE, so
        # its properties are reached through IDispatch; if that fails too,
        # the user opens the editor with Alt+F11.
        try {
            $flags = [Reflection.BindingFlags]
            $vbe = [__ComObject].InvokeMember('VBE', $flags::GetProperty, $null, $excel, $null)
            $window = [__ComObject].InvokeMember('MainWindow', $flags::GetProperty, $null, $vbe, $null)
            [void][__ComObject].InvokeMember('Visible', $flags::SetProperty, $null, $window, @($true))
        } catch {
            Write-Host 'Open the VBA editor with Alt+F11 in the Excel window.'
        }
        [void](Read-Host 'Press Enter here after you have saved and closed the workbook')
        Release-Com $book
    } -Visible
    if ((Get-ProjectProtection $protected) -eq 0) {
        throw 'protected.xlsm is still not locked. Run again and lock the project before saving.'
    }
}
Stop-WindowlessExcel

# -- Sync refusals ----------------------------------------------------------
$cases = @(
    @{ Name = 'protected-project'; Workbook = $protected; Expected = 'VBA project is protected' },
    @{ Name = 'component-type-conflict'; Workbook = $conflict; Expected = 'Component type conflict: M_cPM_TIMEWASTERS' }
)
$results = @()
foreach ($case in $cases) {
    Write-Host ''
    Write-Host "=== $($case.Name) ==="
    $vbomBefore = Get-Vbom
    $hashBefore = (Get-FileHash -LiteralPath $case.Workbook -Algorithm SHA256).Hash.ToLower()
    # No 2>&1: in Windows PowerShell, under Stop, a redirected stderr line would end this script.
    $output = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $tool -Mode Sync `
        -AllowVbomAccess -Workbook $case.Workbook | Out-String
    $exitCode = $LASTEXITCODE
    Write-Host $output.Trim()
    Start-Sleep -Seconds 1
    $vbomAfter = Get-Vbom
    $excelLeft = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
    $bundle = ([regex]::Match($output, '(?m)^Output: (.+?)\s*$')).Groups[1].Value
    $manifest = $null
    if ($bundle -and (Test-Path (Join-Path $bundle 'manifest.json'))) {
        $manifest = Get-Content -Raw (Join-Path $bundle 'manifest.json') | ConvertFrom-Json
    }
    $hashAfter = (Get-FileHash -LiteralPath $case.Workbook -Algorithm SHA256).Hash.ToLower()
    $checks = [ordered]@{
        exit_nonzero      = ($exitCode -ne 0)
        result_fail       = ($output -match 'RESULT=FAIL')
        expected_error    = ($null -ne $manifest -and [string]$manifest.error -like "*$($case.Expected)*")
        not_passed        = ($null -ne $manifest -and -not $manifest.passed)
        cleanup_completed = ($null -ne $manifest -and $manifest.cleanup -eq 'completed')
        input_unchanged   = ($hashBefore -eq $hashAfter)
        vbom_restored     = ($vbomBefore -eq $vbomAfter)
        no_excel_left     = ($excelLeft -eq 0)
    }
    $passed = -not ($checks.Values -contains $false)
    $results += [ordered]@{
        case = $case.Name; workbook = $case.Workbook; expected_error = $case.Expected
        exit_code = $exitCode; vbom_before = $vbomBefore; vbom_after = $vbomAfter
        input_sha256 = $hashBefore; excel_processes_after = $excelLeft
        manifest_error = $(if ($manifest) { $manifest.error } else { $null })
        manifest_cleanup = $(if ($manifest) { $manifest.cleanup } else { $null })
        excel_build = $(if ($manifest) { $manifest.excel_build } else { $null })
        bundle = $bundle; checks = $checks; verdict = $(if ($passed) { 'PASS' } else { 'FAIL' })
    }
    foreach ($check in $checks.GetEnumerator()) {
        Write-Host ("  {0,-18} {1}" -f $check.Key, $(if ($check.Value) { 'ok' } else { 'FAILED' }))
    }
    Write-Host "$($case.Name): $(if ($passed) { 'PASS' } else { 'FAIL' })"
    Stop-WindowlessExcel
}

$verdict = $(if (@($results | Where-Object { $_.verdict -ne 'PASS' }).Count -eq 0) { 'PASS' } else { 'FAIL' })
$stamp = (Get-Date).ToUniversalTime()
$evidence = [ordered]@{
    issue = 138; tool = 'Verify-SyncRefusals.ps1'; commit_sha = $sha
    timestamp_utc = $stamp.ToString('o'); cases = $results; verdict = $verdict
}
$out = Join-Path $WorkDir ('sync-refusals-' + $sha.Substring(0, 7) + '-' + $stamp.ToString('yyyyMMddTHHmmssZ') + '.json')
$evidence | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $out -Encoding UTF8
Write-Host ''
Write-Host "RESULT=$verdict"
Write-Host "Evidence: $out"
if ($verdict -eq 'PASS') { exit 0 } else { exit 1 }
