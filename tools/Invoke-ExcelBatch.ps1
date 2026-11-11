<#
.SYNOPSIS
    Runs tools\Invoke-ExcelRegression.ps1 for several commits in one unattended
    pass, building each red variant from Git (#11, local runner).

.DESCRIPTION
    The manifest (JSON, kept outside the repository) lists the commits to run,
    typically the heads of open pull requests. For each entry the script:
      1. creates a temporary detached worktree at the entry's SHA, so the
         maintainer's own checkout is never touched;
      2. builds the entry's red file, if any, from Git:
           from_ref       - the component exactly as it is at that commit
                            (usually the base the pull request started from);
           find / replace - the component at the entry's SHA with one exact
                            text occurrence replaced;
      3. runs that worktree's own Invoke-ExcelRegression.ps1 with -Repo, which
         still checks the SHA and a clean tree, starts and quits its own Excel
         instance, restores AccessVBOM and writes the JSON evidence record;
      4. removes the worktree and records the verdict.
    It ends with a summary table and a batch record next to the evidence.

    Manifest:

        {
          "schema_version": 1,
          "runs": [
            { "label": "PR #143 (#30)", "sha": "<40-hex>",
              "red": { "component": "src\\classes\\cPerformanceManager.cls",
                       "expected_case": 88, "from_ref": "<40-hex>" } },
            { "label": "PR #145 (#73)", "sha": "<40-hex>" }
          ]
        }

    tools\excel-batch.example.json is a complete example. The whole manifest is
    validated before Excel starts. Missing commits are fetched once.

    Limits: the same as the runner's. There is no timeout, and a modal Excel
    error dialog waits for the maintainer.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File tools\Invoke-ExcelBatch.ps1 `
        -Manifest ..\batch.json -OutDir ..\excel-evidence
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)] [string] $Manifest,
    [Parameter(Mandatory = $true)] [string] $OutDir,
    [string] $Repo
)

Set-StrictMode -Version 2
$ErrorActionPreference = 'Stop'

# The runner's components; a red file may replace exactly one of them.
$Components = @(
    'src\modules\M_cPM_TIMEWASTERS.bas',
    'src\classes\cPerformanceManager.cls',
    'demo\M_DEMO_BUILDER.bas',
    'test\M_cPM_Test.bas'
)
$Sha40 = '^[0-9a-f]{40}$'

if (-not $Repo) { $Repo = Split-Path -Parent $PSScriptRoot }
$Repo = (Resolve-Path $Repo).Path
$Manifest = (Resolve-Path $Manifest).Path
# Resolve against PowerShell's current location, which .NET's GetFullPath ignores.
$OutDir = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutDir)
if (($OutDir.TrimEnd('\') + '\').StartsWith($Repo.TrimEnd('\') + '\', [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "-OutDir must be outside the repository ($Repo), or the evidence would dirty the checkout."
}
$Latin1 = [System.Text.Encoding]::GetEncoding(28591)
$Work = Join-Path ([System.IO.Path]::GetTempPath()) ('cpm-excel-batch-' + [guid]::NewGuid().ToString('N'))

function Get-Field($Object, [string] $Name) {
    if ($null -eq $Object -or $Object -isnot [System.Management.Automation.PSCustomObject]) { return $null }
    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property) { return $null }
    return $property.Value
}

function Invoke-Git([string[]] $Arguments) {
    # No 2>&1: in Windows PowerShell 5.1 a redirected stderr line would stop
    # the script under ErrorActionPreference Stop. The exit code is checked.
    $output = & git -C $Repo @Arguments
    if ($LASTEXITCODE -ne 0) { throw "git $($Arguments -join ' ') failed with exit code $LASTEXITCODE" }
    return $output
}

function Test-Commit([string] $Sha) {
    & git -C $Repo rev-parse --verify --quiet "$Sha^{commit}" | Out-Null
    return ($LASTEXITCODE -eq 0)
}

function Read-Manifest([string] $Path) {
    $data = Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json
    if ((Get-Field $data 'schema_version') -ne 1) { throw "${Path}: schema_version must be 1" }
    $runs = @(Get-Field $data 'runs')
    if ($runs.Count -eq 0 -or $null -eq $runs[0]) { throw "${Path}: runs must list at least one entry" }
    $entries = @()
    $index = 0
    foreach ($run in $runs) {
        $index++
        $where = "${Path}: runs[$index]"
        $sha = [string](Get-Field $run 'sha')
        if ($sha -notmatch $Sha40) { throw "${where}: sha must be a full lowercase commit SHA" }
        $label = [string](Get-Field $run 'label')
        if (-not $label) { $label = $sha.Substring(0, 7) }
        $entry = [ordered]@{ label = $label; sha = $sha; red = $null }
        $red = Get-Field $run 'red'
        if ($null -ne $red) {
            $component = ([string](Get-Field $red 'component')).Replace('/', '\')
            if ($Components -notcontains $component) { throw "${where}: red.component must be one of: $($Components -join ', ')" }
            $case = Get-Field $red 'expected_case'
            if ($case -isnot [int] -or $case -lt 1) { throw "${where}: red.expected_case must be a positive integer" }
            $fromRef = Get-Field $red 'from_ref'
            $find = Get-Field $red 'find'
            $replace = Get-Field $red 'replace'
            if ($null -ne $fromRef) {
                if ($null -ne $find -or $null -ne $replace) { throw "${where}: red takes from_ref or find/replace, not both" }
                if ([string]$fromRef -notmatch $Sha40) { throw "${where}: red.from_ref must be a full lowercase commit SHA" }
            } elseif (-not ($find -is [string] -and $replace -is [string] -and $find -and $find -ne $replace)) {
                throw "${where}: red needs from_ref, or a non-empty find and a different replace"
            }
            $entry.red = [ordered]@{ component = $component; expected_case = [int]$case
                                     from_ref = $fromRef; find = $find; replace = $replace }
        }
        $entries += [pscustomobject]$entry
    }
    return $entries
}

function New-RedFile($Entry, [string] $Worktree) {
    $red = $Entry.red
    $dir = Join-Path $Work ('red-' + $Entry.sha.Substring(0, 7))
    [void](New-Item -ItemType Directory -Path $dir -Force)
    $target = Join-Path $dir (Split-Path -Leaf $red.component)
    if ($null -ne $red.from_ref) {
        # Redirect git's stdout straight to a file, so the bytes stay exact.
        $blob = Join-Path $dir 'blob.tmp'
        $process = Start-Process -FilePath 'git' -WorkingDirectory $Repo -NoNewWindow -Wait -PassThru `
            -ArgumentList @('show', ($red.from_ref + ':' + $red.component.Replace('\', '/'))) `
            -RedirectStandardOutput $blob
        if ($process.ExitCode -ne 0) { throw "git show $($red.from_ref):$($red.component) failed" }
        $text = $Latin1.GetString([System.IO.File]::ReadAllBytes($blob))
        Remove-Item -LiteralPath $blob
        # Git stores the exports normalized; the VBE imports CRLF.
        $text = $text.Replace("`r`n", "`n").Replace("`n", "`r`n")
    } else {
        $text = $Latin1.GetString([System.IO.File]::ReadAllBytes((Join-Path $Worktree $red.component)))
        $count = ([regex]::Matches($text, [regex]::Escape($red.find))).Count
        if ($count -ne 1) { throw "$($Entry.label): red.find occurs $count times in $($red.component), expected once" }
        $text = $text.Replace($red.find, $red.replace)
    }
    [System.IO.File]::WriteAllBytes($target, $Latin1.GetBytes($text))
    return $target
}

function Remove-Worktree([string] $Path) {
    if (Test-Path -LiteralPath $Path) { & git -C $Repo worktree remove --force $Path | Out-Null }
    & git -C $Repo worktree prune | Out-Null
}

# --------------------------------------------------------------------------- #
# Preconditions: the whole manifest, the commits and a closed Excel
# --------------------------------------------------------------------------- #

$entries = @(Read-Manifest $Manifest)
if (Get-Process -Name EXCEL -ErrorAction SilentlyContinue) {
    throw 'Excel is already running. Save your work, close every Excel window, then run again.'
}
$needed = @($entries | ForEach-Object { $_.sha; if ($_.red -and $_.red.from_ref) { $_.red.from_ref } } | Select-Object -Unique)
if (@($needed | Where-Object { -not (Test-Commit $_) }).Count -gt 0) {
    Write-Host 'Fetching origin for commits not present locally...'
    & git -C $Repo fetch --quiet origin
    if ($LASTEXITCODE -ne 0) { throw 'git fetch origin failed.' }
    $missing = @($needed | Where-Object { -not (Test-Commit $_) })
    if ($missing.Count -gt 0) { throw "Commits not found after fetching: $($missing -join ', ')" }
}
[void](New-Item -ItemType Directory -Path $Work)

# --------------------------------------------------------------------------- #
# Runs, one worktree at a time
# --------------------------------------------------------------------------- #

$summary = @()
$index = 0
try {
    foreach ($entry in $entries) {
        $index++
        $worktree = Join-Path $Work ('wt-' + $index)
        Write-Host ''
        Write-Host "=== $index of $($entries.Count): $($entry.label) at $($entry.sha.Substring(0, 7)) ==="
        $verdict = 'NOT RUN'
        try {
            & git -C $Repo worktree add --detach --quiet $worktree $entry.sha
            if ($LASTEXITCODE -ne 0) { throw "git worktree add failed for $($entry.sha)" }
            $runnerArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass',
                            '-File', (Join-Path $worktree 'tools\Invoke-ExcelRegression.ps1'),
                            '-ExpectedSha', $entry.sha, '-Repo', $worktree,
                            '-AllowVbomAccess', '-OutDir', $OutDir)
            if ($entry.red) {
                $redFile = New-RedFile $entry $worktree
                $runnerArgs += @('-RedFile', $redFile, '-RedComponent', $entry.red.component,
                                 '-RedExpectedCase', [string]$entry.red.expected_case)
            }
            & powershell.exe @runnerArgs
            $verdict = $(if ($LASTEXITCODE -eq 0) { 'PASS' } else { "FAIL (runner exit $LASTEXITCODE)" })
        } catch {
            $verdict = "ERROR: $($_.Exception.Message)"
            Write-Host $verdict
        } finally {
            Remove-Worktree $worktree
        }
        $summary += [pscustomobject]([ordered]@{
            label = $entry.label; sha = $entry.sha
            red = $(if ($entry.red) { "case $($entry.red.expected_case)" } else { '-' })
            verdict = $verdict })
    }
} finally {
    if (Test-Path -LiteralPath $Work) { Remove-Item -LiteralPath $Work -Recurse -Force -ErrorAction SilentlyContinue }
}

# --------------------------------------------------------------------------- #
# Summary
# --------------------------------------------------------------------------- #

if (-not (Test-Path $OutDir)) { [void](New-Item -ItemType Directory -Path $OutDir -Force) }
$stamp = (Get-Date).ToUniversalTime().ToString('yyyyMMddTHHmmssZ')
$record = Join-Path $OutDir "cpm-batch-$stamp.json"
[ordered]@{ tool = 'tools/Invoke-ExcelBatch.ps1'; manifest = (Split-Path -Leaf $Manifest)
            finished = (Get-Date).ToUniversalTime().ToString('o'); runs = $summary } |
    ConvertTo-Json -Depth 4 | Set-Content -Path $record -Encoding UTF8

Write-Host ''
Write-Host '=== Summary ==='
$summary | Format-Table -AutoSize | Out-String | Write-Host
Write-Host "Batch record written to $record"
if (@($summary | Where-Object { $_.verdict -ne 'PASS' }).Count -gt 0) { exit 1 } else { exit 0 }
