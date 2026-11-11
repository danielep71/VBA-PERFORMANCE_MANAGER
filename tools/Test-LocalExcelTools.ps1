# Parse all PowerShell tools; execute only pure helpers, never Excel/registry code.
$ErrorActionPreference = 'Stop'
$asts = @{}
foreach ($file in Get-ChildItem $PSScriptRoot -Filter '*.ps1') {
    $tokens = $null; $errors = $null
    $ast = [Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$tokens, [ref]$errors)
    if ($errors.Count) { throw "$($file.Name): $($errors -join '; ')" }
    $asts[$file.Name] = $ast
}
function Import-PureFunction([string] $File, [string] $Name) {
    $nodes = @($asts[$File].FindAll({ param($node)
        $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $Name
    }, $true))
    if ($nodes.Count -ne 1) { throw "Expected one helper: $Name" }
    return $nodes[0].Extent.Text
}
. ([scriptblock]::Create((Import-PureFunction 'Invoke-ExcelRegression.ps1' 'Test-ExampleTranscript')))
. ([scriptblock]::Create((Import-PureFunction 'Manage-PerformanceManagerVBA.ps1' 'Normalize-Code')))
. ([scriptblock]::Create((Import-PureFunction 'Manage-PerformanceManagerVBA.ps1' 'Source-Code')))
. ([scriptblock]::Create((Import-PureFunction 'Manage-PerformanceManagerVBA.ps1' 'Component-Name')))
$names = @('Example_Benchmark_WithStatistics', 'Example_Compare_TwoImplementations',
           'Example_DispatchMatchedBaseline', 'Example_ReadStatus_Diagnostics')
$good = "PASS`n"
foreach ($name in $names) { $good += "BEGIN $name`nordinary output`nEND $name`n" }
if (-not (Test-ExampleTranscript $good)) { throw 'Complete transcript rejected' }
$bad = @($good.Replace("PASS`n", "FAIL`n"),
         $good.Replace("END $($names[2])`n", ''),
         ($good + "END $($names[1])`n"),
         ($good + "Error 13 - fixture`n"),
         $good.Replace("BEGIN $($names[0])", "END $($names[0])").Replace("END $($names[0])`nordinary", "BEGIN_WRONG`nordinary"))
foreach ($text in $bad) { if (Test-ExampleTranscript $text) { throw 'Invalid transcript accepted' } }
$temp = Join-Path ([IO.Path]::GetTempPath()) ('cpm-code-' + [guid]::NewGuid() + '.cls')
try {
    $source = "VERSION 1.0 CLASS`r`nBEGIN`r`nEND`r`nAttribute VB_Name = `"Fixture`"`r`nAttribute VB_Exposed = False`r`nOption Explicit`r`nPublic Sub Example()`r`nEnd Sub`r`n"
    [IO.File]::WriteAllText($temp, $source)
    $expected = "Option Explicit`nPublic Sub Example()`nEnd Sub"
    if ((Source-Code $temp) -cne $expected) { throw 'Exported class metadata normalization failed' }
    if ((Component-Name $temp) -cne 'Fixture') { throw 'Component name not read from VB_Name' }
    # The file name and the VB_Name differ in case; sync must use the VB_Name
    $timeWasters = Join-Path (Split-Path -Parent $PSScriptRoot) 'src\modules\M_cPM_TIMEWASTERS.bas'
    if ((Component-Name $timeWasters) -cne 'M_cPM_TimeWasters') { throw 'VB_Name case not preserved' }
    if ((Normalize-Code "Option Explicit  `r`nPublic Sub Example()`r`nEnd Sub`r`n") -cne $expected) { throw 'Line-ending normalization failed' }
    if ((Normalize-Code ($expected.Replace('Example','Changed'))) -ceq (Source-Code $temp)) { throw 'Changed code accepted' }
    [IO.File]::WriteAllText($temp, 'no exported identity')
    $rejected = $false
    try { [void](Source-Code $temp) } catch { $rejected = $true }
    if (-not $rejected) { throw 'Missing exported identity accepted' }
    $rejected = $false
    try { [void](Component-Name $temp) } catch { $rejected = $true }
    if (-not $rejected) { throw 'Missing component name accepted' }
} finally { Remove-Item -LiteralPath $temp -Force }
Write-Host 'PASS: PowerShell parsing, transcript failures and code readback fixtures'
exit 0
