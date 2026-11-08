# Exercise the actual verdict/evidence/exit section with synthetic run results.
# No Excel or Windows registry access is performed.
$ErrorActionPreference = 'Stop'
$text = Get-Content (Join-Path $PSScriptRoot 'Invoke-ExcelRegression.ps1') -Raw
$marker = '$source = $runs[0]'
$offset = $text.IndexOf($marker, [StringComparison]::Ordinal)
if ($offset -lt 0) { throw 'Runner verdict section not found' }
$tail = $text.Substring($offset)
$shellPath = (Get-Process -Id $PID).Path
$temp = Join-Path ([IO.Path]::GetTempPath()) ('cpm-verdict-' + [guid]::NewGuid())
[void](New-Item -ItemType Directory $temp)
$scenarios = @(
    @{ name = 'clean red'; cleanup = 'completed'; source = 'completed'; expected = 0 },
    @{ name = 'forced red'; cleanup = 'forced: Excel still running'; source = 'completed'; expected = 1 },
    @{ name = 'null red'; cleanup = $null; source = 'completed'; expected = 1 },
    @{ name = 'missing red'; source = 'completed'; expected = 1 },
    @{ name = 'forced source'; cleanup = 'completed'; source = 'forced'; expected = 1 },
    @{ name = 'wrong failed case'; cleanup = 'completed'; source = 'completed'; wrongCase = $true; expected = 1 },
    @{ name = 'clean source only'; source = 'completed'; sourceOnly = $true; expected = 0 },
    @{ name = 'forced source only'; source = 'forced'; sourceOnly = $true; expected = 1 }
)
try {
    foreach ($scenario in $scenarios) {
        $source = [ordered]@{
            run = 'source'; state = 'completed'; cases = 86; assertions = 714
            failures = 0; failed_cases = @(); cleanup = $scenario.source
            excel = $null; error = $null; finished = '2026-11-08T00:00:00Z'
        }
        $red = [ordered]@{
            run = 'red check'; state = 'completed'; cases = 86; assertions = 714
            failures = 1; failed_cases = @(@{ case = $(if ($scenario.wrongCase) { 85 } else { 86 }); name = 'fixture' })
            excel = $null; error = $null; finished = '2026-11-08T00:00:01Z'
        }
        if ($scenario.ContainsKey('cleanup')) { $red.cleanup = $scenario.cleanup }
        $fixtureRuns = @($source)
        if (-not $scenario.sourceOnly) { $fixtureRuns += $red }
        $json = ConvertTo-Json -InputObject $fixtureRuns -Depth 6 -Compress
        $out = Join-Path $temp $scenario.name
        $redFileValue = if ($scenario.sourceOnly) { "" } else { "red.cls" }
        $header = @"
`$runs = ConvertFrom-Json '$($json.Replace("'", "''"))'
`$sha = '0123456789012345678901234567890123456789'
`$branch = 'test-fixture'
`$expectedCases = 86
`$RedFile = '$redFileValue'
`$RedExpectedCase = 86
`$vbomChanged = `$false
`$OutDir = '$($out.Replace("'", "''"))'
"@
        $script = Join-Path $temp 'fixture.ps1'
        Set-Content $script ($header + "`n" + $tail) -Encoding utf8
        $output = & $shellPath -NoProfile -File $script 2>&1
        if ($LASTEXITCODE -ne $scenario.expected) {
            throw "$($scenario.name): wrong exit $LASTEXITCODE; expected $($scenario.expected). $output"
        }
        $files = @(Get-ChildItem $out -Filter '*.json')
        if ($files.Count -ne 1) { throw "$($scenario.name): missing evidence" }
        $record = Get-Content $files[0].FullName -Raw | ConvertFrom-Json
        if ($record.result.cleanup_state -ne $scenario.source -or
            $record.details.runs[0].cleanup -ne $scenario.source) {
            throw "$($scenario.name): source cleanup evidence lost"
        }
        if (-not $scenario.sourceOnly) {
            if ($record.details.runs[1].cleanup -ne $scenario.cleanup) {
                throw "$($scenario.name): red cleanup evidence lost"
            }
            $expectedRed = if ($scenario.cleanup -eq 'completed' -and -not $scenario.wrongCase) {
                'AS EXPECTED: only case 86 fails'
            } else { 'NOT AS EXPECTED' }
            if ($record.details.verdicts.red_check -ne $expectedRed) {
                throw "$($scenario.name): wrong red verdict"
            }
        }
        Write-Host "PASS: $($scenario.name)"
    }
} finally {
    Remove-Item $temp -Recurse -Force
}
