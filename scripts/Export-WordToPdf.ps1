param(
    [Parameter(Mandatory = $false)]
    [string]$InputPath = "docs"
)

$ErrorActionPreference = "Stop"

function Release-ComObject {
    param([object]$ComObject)
    if ($null -ne $ComObject) {
        try {
            [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ComObject)
        } catch {
        }
    }
}

$resolvedInput = Resolve-Path -LiteralPath $InputPath -ErrorAction Stop
$item = Get-Item -LiteralPath $resolvedInput

if ($item.PSIsContainer) {
    $docxFiles = @(Get-ChildItem -LiteralPath $item.FullName -Filter *.docx -File -Recurse |
        Where-Object { $_.Name -notmatch '^~\$' })
} else {
    if ($item.Extension -ne ".docx") {
        throw "Input file is not a .docx file: $($item.FullName)"
    }
    $docxFiles = @($item)
}

if ($docxFiles.Count -eq 0) {
    Write-Host "No .docx files found under: $($item.FullName)"
    exit 0
}

Write-Host "Found $($docxFiles.Count) Word document(s)."

$word = $null

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0

    try { $word.AutomationSecurity = 3 } catch {}

    Write-Host "Microsoft Word version: $($word.Version)"

    foreach ($docx in $docxFiles) {
        $document = $null
        try {
            $pdfPath = [System.IO.Path]::ChangeExtension($docx.FullName, ".pdf")

            Write-Host ""
            Write-Host "Rendering:"
            Write-Host "  DOCX: $($docx.FullName)"
            Write-Host "  PDF : $pdfPath"

            $document = $word.Documents.Open($docx.FullName, $false, $true)

            try { $document.Repaginate() } catch {}
            try {
                foreach ($field in $document.Fields) {
                    try { [void]$field.Update() } catch {}
                    Release-ComObject $field
                }
            } catch {}

            $document.ExportAsFixedFormat($pdfPath, 17)

            if (-not (Test-Path -LiteralPath $pdfPath)) {
                throw "Word did not create the expected PDF: $pdfPath"
            }

            $pdf = Get-Item -LiteralPath $pdfPath
            if ($pdf.Length -le 0) {
                throw "Generated PDF is empty: $pdfPath"
            }

            $pages = $null
            try { $pages = $document.ComputeStatistics(2) } catch {}

            if ($null -ne $pages) {
                Write-Host "  Pages: $pages"
            }
            Write-Host "  Size : $([Math]::Round($pdf.Length / 1KB, 1)) KB"
            Write-Host "  Result: OK"
        }
        finally {
            if ($null -ne $document) {
                try { $document.Close($false) } catch {}
                Release-ComObject $document
                $document = $null
            }
        }
    }
}
finally {
    if ($null -ne $word) {
        try { $word.Quit() } catch {}
        Release-ComObject $word
        $word = $null
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Write-Host ""
Write-Host "All Word documents rendered successfully."
