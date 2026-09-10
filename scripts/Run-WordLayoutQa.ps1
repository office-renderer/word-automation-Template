param(
    [string]$InputPath = ".",
    [string]$ReportRoot = "_qa"
)

$ErrorActionPreference = "Stop"

$input = (Resolve-Path -LiteralPath $InputPath -ErrorAction Stop).Path
New-Item -ItemType Directory -Path $ReportRoot -Force | Out-Null

$item = Get-Item -LiteralPath $input

if ($item.PSIsContainer) {
    $root = $item.FullName
    $docxFiles = @(Get-ChildItem -LiteralPath $root -Filter *.docx -File -Recurse |
        Where-Object {
            $_.Name -notmatch '^~\$' -and
            $_.FullName -notmatch '[\\/]\.git[\\/]'
        })
} else {
    $root = Split-Path -Parent $item.FullName
    $docxFiles = @($item)
}

if ($docxFiles.Count -eq 0) {
    Write-Host "No DOCX files found for QA."
    exit 0
}

$summary = New-Object System.Collections.Generic.List[string]
$summary.Add("# Word Layout QA Summary")
$summary.Add("")
$summary.Add("| Document | Status | Report |")
$summary.Add("|---|---|---|")

foreach ($docx in $docxFiles) {
    $relative = [System.IO.Path]::GetRelativePath($root, $docx.FullName)
    $safe = ($relative -replace '[\\/:*?"<>|]', '__')
    $report = Join-Path $ReportRoot ($safe + ".qa.md")

    Write-Host ""
    Write-Host "QA: $relative"
    & "$PSScriptRoot\Inspect-WordLayout.ps1" -InputPath $docx.FullName -ReportPath $report

    $status = "UNKNOWN"
    if (Test-Path -LiteralPath $report) {
        $match = Select-String -LiteralPath $report -Pattern '^- Status: \*\*(.+)\*\*$' | Select-Object -First 1
        if ($null -ne $match) {
            $status = $match.Matches[0].Groups[1].Value
        }
    }

    $summary.Add("| $relative | $status | $report |")
}

$summaryPath = Join-Path $ReportRoot "summary.md"
[System.IO.File]::WriteAllLines(
    [System.IO.Path]::GetFullPath($summaryPath),
    $summary,
    [System.Text.UTF8Encoding]::new($false)
)

Write-Host ""
Write-Host "QA summary: $summaryPath"
