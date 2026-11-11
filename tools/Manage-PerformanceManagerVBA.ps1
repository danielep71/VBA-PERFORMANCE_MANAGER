<#
.SYNOPSIS
Copy-and-sync or export a Performance Manager .xlsm workbook.
.DESCRIPTION
Supervised Windows desktop Excel tool. Never overwrites the input workbook or
exports into the source checkout. Opening macros and events are disabled.
Sync imports the six repository modules into a fresh copy, compiles and checks
code readback before and after saving. Export never writes to the workbook.
No release certification or security audit is implied. Modal Excel dialogs
require supervision; this is not an unattended CI runner.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)] [ValidateSet('Sync','Export')] [string] $Mode,
    [string] $Workbook,
    [switch] $AllowVbomAccess
)
Set-StrictMode -Version 2
$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$components = @('src\modules\M_cPM_TIMEWASTERS.bas', 'src\classes\cPerformanceManager.cls',
    'demo\M_DEMO_BUILDER.bas', 'demo\M_cPM_DEMO.bas', 'demo\M_cPM_USAGE_EXAMPLES.bas',
    'test\M_cPM_Test.bas')
$vbomKey = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security'
$excel = $null; $book = $null; $project = $null; $control = $null; $excelPid = 0
$bundle = $null; $vbomTouched = $false; $vbomExisted = $false; $vbomBefore = $null
$exitCode = 1
$record = [ordered]@{ mode = $Mode; state = 'not started'; commit_sha = $null
    input_sha256 = $null; original_unchanged = $false; files = @(); components = @()
    compiled = $false; reopened_verified = $false; case_normalized = @(); cleanup = 'not started'; error = $null }
function Release-Com($obj) {
    if ($null -ne $obj) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($obj) }
}
function Normalize-Code([string] $code) {
    $lines = $code -split '\r?\n'
    return (($lines | ForEach-Object { $_.TrimEnd() }) -join "`n").Trim()
}
function Source-Code([string] $path) {
    $text = [IO.File]::ReadAllText($path)
    $match = [regex]::Match($text, '(?m)^Attribute VB_Name = "[^"]+"\r?\n')
    if (-not $match.Success) { throw "Missing VB_Name in $path" }
    $text = $text.Substring($match.Index + $match.Length)
    $text = [regex]::Replace($text, '(?m)^Attribute .*(\r?\n|$)', '')
    return Normalize-Code $text
}
function Component-Name([string] $path) {
    # VBA names a component by its exported VB_Name, which can differ from the
    # file name in case (M_cPM_TIMEWASTERS.bas holds M_cPM_TimeWasters).
    $text = [IO.File]::ReadAllText($path)
    $match = [regex]::Match($text, '(?m)^Attribute VB_Name = "([^"]+)"\r?$')
    if (-not $match.Success) { throw "Missing VB_Name in $path" }
    return $match.Groups[1].Value
}
function Compare-Readback([string] $actual, [string] $expected) {
    # The VBE keeps one spelling per identifier across the whole project, so
    # importing a later module can recase identifiers in an earlier one (a
    # local v becomes V). VBA is case-insensitive: such a difference is
    # reported as 'case', anything else as 'differs'.
    if ([string]::Equals($actual, $expected, [StringComparison]::Ordinal)) { return 'exact' }
    if ([string]::Equals($actual, $expected, [StringComparison]::OrdinalIgnoreCase)) { return 'case' }
    return 'differs'
}
function Verify-Modules($project) {
    foreach ($relative in $components) {
        $path = Join-Path $repo $relative
        $name = Component-Name $path
        $part = $null; $code = $null
        try {
            $part = $project.VBComponents.Item($name)
            $code = $part.CodeModule
            $actual = if ($code.CountOfLines -gt 0) { [string]$code.Lines(1, $code.CountOfLines) } else { '' }
            $outcome = Compare-Readback (Normalize-Code $actual) (Source-Code $path)
            if ($outcome -eq 'differs') { throw "Code readback differs: $name" }
            if ($outcome -eq 'case' -and $record.case_normalized -notcontains $name) {
                $record.case_normalized += $name
            }
        } finally { Release-Com $code; Release-Com $part }
    }
}
try {
    if (Get-Process -Name EXCEL -ErrorAction SilentlyContinue) { throw 'Save your work and close Excel first.' }
    if (-not $Workbook) {
        Add-Type -AssemblyName System.Windows.Forms
        $dialog = New-Object System.Windows.Forms.OpenFileDialog
        try {
            $dialog.Filter = 'Excel macro-enabled workbook (*.xlsm)|*.xlsm'
            $dialog.Title = "$Mode Performance Manager VBA"
            if ($dialog.ShowDialog() -ne 'OK') { throw 'Selection cancelled; no workbook changed.' }
            $Workbook = $dialog.FileName
        } finally { $dialog.Dispose() }
    }
    $Workbook = (Resolve-Path -LiteralPath $Workbook).Path
    if ([IO.Path]::GetExtension($Workbook) -ine '.xlsm') { throw 'Only .xlsm workbooks are supported.' }
    if ($Mode -eq 'Sync') {
        $sha = (& git -C $repo rev-parse HEAD | Out-String).Trim()
        if ($LASTEXITCODE -ne 0 -or $sha -notmatch '^[a-f0-9]{40}$') { throw 'Sync requires a Git checkout.' }
        $dirty = (& git -C $repo status --porcelain | Out-String).Trim()
        if ($LASTEXITCODE -ne 0 -or $dirty) { throw 'Sync requires a clean checkout.' }
        $record.commit_sha = $sha
        foreach ($relative in $components) {
            $path = Join-Path $repo $relative
            if (-not (Test-Path -LiteralPath $path)) { throw "Missing source: $relative" }
            $record.files += [ordered]@{ source = $relative; sha256 = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLower() }
        }
    }
    $record.input_sha256 = (Get-FileHash -LiteralPath $Workbook -Algorithm SHA256).Hash.ToLower()
    $bundle = Join-Path (Split-Path -Parent $repo) ('cpm-' + $Mode.ToLower() + '-' + [Guid]::NewGuid().ToString('N'))
    [void](New-Item -ItemType Directory -Path $bundle)
    $openPath = $Workbook
    if ($Mode -eq 'Sync') {
        $openPath = Join-Path $bundle ([IO.Path]::GetFileName($Workbook))
        Copy-Item -LiteralPath $Workbook -Destination $openPath
    }
    if ($AllowVbomAccess) {
        if (-not (Test-Path $vbomKey)) { [void](New-Item -Path $vbomKey -Force) }
        $before = Get-ItemProperty -Path $vbomKey -Name AccessVBOM -ErrorAction SilentlyContinue
        if ($null -ne $before) { $vbomExisted = $true; $vbomBefore = $before.AccessVBOM }
        $vbomTouched = $true
        Set-ItemProperty -Path $vbomKey -Name AccessVBOM -Value 1 -Type DWord
    }
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class CpmWorkbookWin32 {
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
'@
    $excel = New-Object -ComObject Excel.Application
    [uint32] $pidValue = 0
    [void][CpmWorkbookWin32]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$pidValue)
    $excelPid = $pidValue
    $excel.AutomationSecurity = 3
    $excel.EnableEvents = $false
    $excel.AskToUpdateLinks = $false
    $excel.Visible = $true
    $record.excel_version = $excel.Version
    $record.excel_build = $excel.Build
    $book = $excel.Workbooks.Open($openPath, 0, ($Mode -eq 'Export'))
    $project = $book.VBProject
    if ($project.Protection -ne 0) { throw 'VBA project is protected. No password removal is attempted.' }
    $record.state = 'processing'
    if ($Mode -eq 'Sync') {
        foreach ($relative in $components) {
            $path = Join-Path $repo $relative
            $name = Component-Name $path
            $expectedType = if ($path.EndsWith('.cls')) { 2 } else { 1 }
            $part = $null
            try {
                for ($i = $project.VBComponents.Count; $i -ge 1; $i--) {
                    $candidate = $project.VBComponents.Item($i)
                    try {
                        if ($candidate.Name -ieq $name) {
                            if ($candidate.Type -ne $expectedType) { throw "Component type conflict: $name" }
                            $project.VBComponents.Remove($candidate)
                        }
                    } finally { Release-Com $candidate }
                }
                $part = $project.VBComponents.Import($path)
                if ($part.Name -cne $name) { throw "Imported component renamed: $name" }
            } finally { Release-Com $part }
        }
        Verify-Modules $project
        $control = $excel.VBE.CommandBars.FindControl([Reflection.Missing]::Value, 578)
        if ($null -eq $control) { throw 'Compile command unavailable.' }
        if ($control.Enabled) { $control.Execute() }
        if ($control.Enabled) { throw 'VBA compile failed.' }
        $record.compiled = $true
        $book.Save()
        Release-Com $control; $control = $null
        Release-Com $project; $project = $null
        $book.Close($false); Release-Com $book; $book = $null
        $book = $excel.Workbooks.Open($openPath, 0, $true)
        $project = $book.VBProject
        Verify-Modules $project
        $record.reopened_verified = $true
        $record.output_file = [IO.Path]::GetFileName($openPath)
        $record.output_sha256 = (Get-FileHash -LiteralPath $openPath -Algorithm SHA256).Hash.ToLower()
    } else {
        for ($i = 1; $i -le $project.VBComponents.Count; $i++) {
            $part = $project.VBComponents.Item($i)
            try {
                $extension = switch ($part.Type) { 1 { '.bas' } 2 { '.cls' } 3 { '.frm' } 100 { '.cls' } default { throw 'Unsupported VBA component type.' } }
                if ($part.Name -notmatch '^[A-Za-z_][A-Za-z0-9_]*$') { throw 'Unsafe component filename.' }
                $file = Join-Path $bundle ($part.Name + $extension)
                $part.Export($file)
                $record.components += [ordered]@{ name = $part.Name; type = [int]$part.Type; file = [IO.Path]::GetFileName($file) }
            } finally { Release-Com $part }
        }
        foreach ($file in Get-ChildItem -LiteralPath $bundle -File) {
            $record.files += [ordered]@{ file = $file.Name; sha256 = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash.ToLower() }
        }
    }
    $record.state = 'completed'
} catch { $record.error = $_.Exception.Message; Write-Host "FAIL: $($record.error)" }
finally {
    try {
        Release-Com $control; Release-Com $project
        $control = $null; $project = $null
        if ($null -ne $book) { $book.Close($false) }
        if ($null -ne $excel) { $excel.Quit() }
        Release-Com $book; Release-Com $excel
        $book = $null; $excel = $null
        for ($i = 0; $i -lt 2; $i++) { [GC]::Collect(); [GC]::WaitForPendingFinalizers() }
        $record.cleanup = 'completed'
    } catch { $record.cleanup = 'failed: ' + $_.Exception.Message }
    if ($excelPid) {
        $deadline = (Get-Date).AddSeconds(15)
        while ((Get-Process -Id $excelPid -ErrorAction SilentlyContinue) -and (Get-Date) -lt $deadline) { Start-Sleep -Milliseconds 250 }
        if (Get-Process -Id $excelPid -ErrorAction SilentlyContinue) {
            Stop-Process -Id $excelPid -Force -ErrorAction SilentlyContinue
            $record.cleanup = 'forced Excel termination; not PASS'
        }
    }
    if ($vbomTouched) {
        try {
            if ($vbomExisted) { Set-ItemProperty -Path $vbomKey -Name AccessVBOM -Value $vbomBefore -Type DWord }
            else { Remove-ItemProperty -Path $vbomKey -Name AccessVBOM -ErrorAction Stop }
        } catch { $record.cleanup = 'failed to restore AccessVBOM: ' + $_.Exception.Message }
    }
    if ($record.input_sha256) {
        try { $record.original_unchanged = ((Get-FileHash -LiteralPath $Workbook -Algorithm SHA256).Hash.ToLower() -eq $record.input_sha256) }
        catch { $record.error = 'Cannot verify original workbook hash: ' + $_.Exception.Message }
    }
    $record.passed = ($record.state -eq 'completed' -and $record.cleanup -eq 'completed' -and $record.original_unchanged -and -not $record.error)
    $record.timestamp_utc = (Get-Date).ToUniversalTime().ToString('o')
    if ($bundle) {
        $record | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $bundle 'manifest.json') -Encoding UTF8
        Write-Host "Output: $bundle"
        try { Invoke-Item -LiteralPath $bundle } catch { }
    }
    if ($record.passed) { $exitCode = 0; Write-Host 'RESULT=PASS' } else { Write-Host 'RESULT=FAIL; do not use an incomplete output workbook.' }
}
exit $exitCode
