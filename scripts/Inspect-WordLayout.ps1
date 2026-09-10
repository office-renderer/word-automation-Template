param(
    [Parameter(Mandatory = $true)]
    [string]$InputPath,
    [string]$ReportPath = "_qa/qa-report.md"
)

$ErrorActionPreference = "Stop"

function Release-ComObject {
    param([object]$ComObject)
    if ($null -ne $ComObject) {
        try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ComObject) } catch {}
    }
}

$inputFull = (Resolve-Path -LiteralPath $InputPath).Path
$reportFull = [System.IO.Path]::GetFullPath((Join-Path $PWD $ReportPath))
New-Item -ItemType Directory -Path (Split-Path -Parent $reportFull) -Force | Out-Null

$findings = New-Object System.Collections.Generic.List[object]

function Add-Finding {
    param(
        [string]$Severity,
        [string]$Rule,
        [string]$Location,
        [string]$Message
    )
    $findings.Add([pscustomobject]@{
        Severity = $Severity
        Rule = $Rule
        Location = $Location
        Message = $Message
    })
}

$word = $null
$doc = $null

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $doc = $word.Documents.Open($inputFull, $false, $true)
    $doc.Repaginate()

    $pageCount = $doc.ComputeStatistics(2)
    Write-Host "QA inspecting $pageCount page(s)."

    $section = $doc.Sections.Item(1)
    $pageWidth = [double]$section.PageSetup.PageWidth
    $pageHeight = [double]$section.PageSetup.PageHeight
    $leftMargin = [double]$section.PageSetup.LeftMargin
    $rightMargin = [double]$section.PageSetup.RightMargin
    $usableWidth = $pageWidth - $leftMargin - $rightMargin
    Release-ComObject $section

    # Rule 1: blank pages based on actual Word pagination.
    for ($p = 1; $p -le $pageCount; $p++) {
        $start = $doc.GoTo(1, 1, $p)
        $startPos = [int]$start.Start
        Release-ComObject $start

        if ($p -lt $pageCount) {
            $next = $doc.GoTo(1, 1, $p + 1)
            $endPos = [int]$next.Start - 1
            Release-ComObject $next
        } else {
            $contentRange = $doc.Content
            $endPos = [int]$contentRange.End
            Release-ComObject $contentRange
        }

        if ($endPos -lt $startPos) { $endPos = $startPos }
        $pageRange = $doc.Range($startPos, $endPos)
        $pageText = [string]$pageRange.Text
        $visible = ($pageText -replace "[\s\r\n\f\a\v]+", "")

        $shapeOnPage = $false
        $shapeCount = $doc.Shapes.Count
        for ($si = 1; $si -le $shapeCount; $si++) {
            $sh = $doc.Shapes.Item($si)
            try {
                $anchor = $sh.Anchor
                $anchorPage = [int]$anchor.Information(3)
                Release-ComObject $anchor
                if ($anchorPage -eq $p) { $shapeOnPage = $true }
            } catch {}
            Release-ComObject $sh
        }

        if ([string]::IsNullOrWhiteSpace($visible) -and -not $shapeOnPage) {
            Add-Finding "ERROR" "BLANK_PAGE" "Page $p" "Word pagination contains an apparently blank page."
        }

        Release-ComObject $pageRange
    }

    # Rule 2: tables wider than the usable text area.
    $tableCount = $doc.Tables.Count
    for ($tableIndex = 1; $tableIndex -le $tableCount; $tableIndex++) {
        $tbl = $doc.Tables.Item($tableIndex)
        $totalWidth = 0.0
        $colCount = $tbl.Columns.Count

        for ($ci = 1; $ci -le $colCount; $ci++) {
            $col = $tbl.Columns.Item($ci)
            try { $totalWidth += [double]$col.Width } catch {}
            Release-ComObject $col
        }

        $page = "?"
        try {
            $tableRange = $tbl.Range
            $page = [int]$tableRange.Information(3)
            Release-ComObject $tableRange
        } catch {}

        if ($totalWidth -gt ($usableWidth + 2)) {
            Add-Finding "ERROR" "TABLE_WIDTH" "Table $tableIndex / Page $page" ("Table width {0:N1} pt exceeds usable text width {1:N1} pt." -f $totalWidth, $usableWidth)
        }

        Release-ComObject $tbl
    }

    # Rule 3: floating shapes outside physical page bounds.
    $shapeCount = $doc.Shapes.Count
    for ($shapeIndex = 1; $shapeIndex -le $shapeCount; $shapeIndex++) {
        $sh = $doc.Shapes.Item($shapeIndex)
        $page = "?"
        try {
            $anchor = $sh.Anchor
            $page = [int]$anchor.Information(3)
            Release-ComObject $anchor
        } catch {}

        $left = [double]$sh.Left
        $top = [double]$sh.Top
        $width = [double]$sh.Width
        $height = [double]$sh.Height

        $outside = ($left -lt 0) -or ($top -lt 0) -or (($left + $width) -gt ($pageWidth + 1)) -or (($top + $height) -gt ($pageHeight + 1))
        if ($outside) {
            Add-Finding "ERROR" "SHAPE_PAGE_BOUNDS" "Shape $shapeIndex / Page $page" ("Floating object bounds left={0:N1}, top={1:N1}, width={2:N1}, height={3:N1} exceed page {4:N1} x {5:N1} pt." -f $left, $top, $width, $height, $pageWidth, $pageHeight)
        }

        Release-ComObject $sh
    }

    # Rule 4: suspicious paragraph font sizes.
    $paragraphCount = $doc.Paragraphs.Count
    for ($paragraphIndex = 1; $paragraphIndex -le $paragraphCount; $paragraphIndex++) {
        $para = $doc.Paragraphs.Item($paragraphIndex)
        $range = $para.Range
        $text = ([string]$range.Text -replace "[\r\a\f]+", "").Trim()

        if (-not [string]::IsNullOrWhiteSpace($text)) {
            $fontSize = [double]$range.Font.Size
            $page = "?"
            try { $page = [int]$range.Information(3) } catch {}

            if ($fontSize -gt 0 -and $fontSize -lt 8) {
                Add-Finding "WARNING" "TINY_FONT" "Paragraph $paragraphIndex / Page $page" ("Paragraph uses suspiciously small font size: {0:N1} pt." -f $fontSize)
            }
        }

        Release-ComObject $range
        Release-ComObject $para
    }

    # Rule 5: very short final visual line in a multi-line paragraph.
    # Uses Word's real line layout through Selection navigation.
    $selection = $word.Selection
    for ($paragraphIndex = 1; $paragraphIndex -le $paragraphCount; $paragraphIndex++) {
        $para = $doc.Paragraphs.Item($paragraphIndex)
        $range = $para.Range
        $text = ([string]$range.Text -replace "[\r\a\f]+", "").Trim()

        $inTable = $false
        try { $inTable = [bool]$range.Information(12) } catch {}

        if (-not $inTable -and -not [string]::IsNullOrWhiteSpace($text)) {
            $lineCount = 0
            try { $lineCount = [int]$range.ComputeStatistics(1) } catch {}

            if ($lineCount -ge 2) {
                try {
                    $selection.SetRange([int]$range.Start, [int]$range.Start)
                    [void]$selection.MoveDown(5, $lineCount - 1, 0)
                    [void]$selection.HomeKey(5, 0)
                    [void]$selection.EndKey(5, 1)

                    $lastLine = ([string]$selection.Text -replace "[\s\r\n\f\a\v]+", "")
                    if ($lastLine.Length -gt 0 -and $lastLine.Length -le 4) {
                        $page = "?"
                        try { $page = [int]$range.Information(3) } catch {}
                        Add-Finding "WARNING" "SHORT_FINAL_LINE" "Paragraph $paragraphIndex / Page $page" ("Final visual line contains only {0} non-whitespace character(s): '{1}'." -f $lastLine.Length, $lastLine)
                    }
                } catch {
                    Write-Host "Short-line heuristic skipped paragraph ${paragraphIndex}: $($_.Exception.Message)"
                }
            }
        }

        Release-ComObject $range
        Release-ComObject $para
    }
    Release-ComObject $selection

    $errors = @($findings | Where-Object Severity -eq "ERROR").Count
    $warnings = @($findings | Where-Object Severity -eq "WARNING").Count
    $status = if ($errors -gt 0) { "FAIL" } elseif ($warnings -gt 0) { "WARN" } else { "PASS" }

    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add("# Word Layout QA Report")
    $lines.Add("")
    $lines.Add("- File: " + $InputPath)
    $lines.Add("- Word pages: " + $pageCount)
    $lines.Add("- Status: **" + $status + "**")
    $lines.Add("- Errors: " + $errors)
    $lines.Add("- Warnings: " + $warnings)
    $lines.Add("")
    $lines.Add("## Findings")
    $lines.Add("")

    if ($findings.Count -eq 0) {
        $lines.Add("No machine-detectable layout issues were found.")
    } else {
        $lines.Add("| Severity | Rule | Location | Finding |")
        $lines.Add("|---|---|---|---|")
        foreach ($f in $findings) {
            $msg = $f.Message.Replace("|", "\|")
            $lines.Add("| " + $f.Severity + " | " + $f.Rule + " | " + $f.Location + " | " + $msg + " |")
        }
    }

    $lines.Add("")
    $lines.Add("## Notes")
    $lines.Add("")
    $lines.Add("This report is structural QA from Microsoft Word COM. Visual QA is still required for aesthetic spacing, typography balance, image quality, and other issues that are difficult to infer from document structure alone.")

    [System.IO.File]::WriteAllLines($reportFull, $lines, [System.Text.UTF8Encoding]::new($false))

    Write-Host "QA_STATUS=$status"
    Write-Host "QA_ERRORS=$errors"
    Write-Host "QA_WARNINGS=$warnings"
    Write-Host "QA report: $reportFull"
}
finally {
    if ($null -ne $doc) {
        try { $doc.Close($false) } catch {}
    }
    if ($null -ne $word) {
        try { $word.Quit() } catch {}
    }

    Release-ComObject $doc
    Release-ComObject $word

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
