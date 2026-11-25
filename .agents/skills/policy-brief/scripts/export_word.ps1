param([Parameter(Mandatory=$true)][string]$BuildDir)
# Windows + installed Microsoft Word. Does not change global security settings.
$ErrorActionPreference = 'Stop'
$buildPath = (Resolve-Path -LiteralPath $BuildDir).Path
$fileNames = @('04-报送清稿.docx', '04-报送核对稿.docx')
$logPath = Join-Path $buildPath 'word-render-check.json'
if (Test-Path -LiteralPath $logPath) { throw 'Refusing to overwrite the render check. Use a new version directory.' }
foreach ($fileName in $fileNames) {
    $source = Join-Path $buildPath $fileName
    $pdf = Join-Path $buildPath ([IO.Path]::GetFileNameWithoutExtension($fileName) + '.pdf')
    if (-not (Test-Path -LiteralPath $source -PathType Leaf)) { throw "Missing DOCX: $fileName" }
    if (Test-Path -LiteralPath $pdf) { throw 'Refusing to overwrite a PDF. Use a new version directory.' }
}
$word = $null
$results = @()
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.AutomationSecurity = 3
    foreach ($fileName in $fileNames) {
        $source = Join-Path $buildPath $fileName
        $pdf = Join-Path $buildPath ([IO.Path]::GetFileNameWithoutExtension($fileName) + '.pdf')
        $before = (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
        $document = $null
        try {
            $document = $word.Documents.Open($source, $false, $true, $false)
            if (-not $document.ReadOnly) { throw 'Word did not open the file read-only.' }
            foreach ($field in $document.Fields) {
                if ($field.Type -in @(33,26,66)) { [void]$field.Update() }
            }
            foreach ($section in $document.Sections) {
                foreach ($kind in @(1,2,3)) {
                    foreach ($story in @($section.Footers.Item($kind), $section.Headers.Item($kind))) {
                        if ($story.Exists) {
                            foreach ($field in $story.Range.Fields) {
                                if ($field.Type -in @(33,26,66)) { [void]$field.Update() }
                            }
                        }
                    }
                }
            }
            $document.Repaginate()
            $document.ExportAsFixedFormat($pdf, 17, $false, 0, 0, 1, 1, 0, $false)
            $results += [ordered]@{
                file = $fileName
                pages = $document.ComputeStatistics(2)
                readonly = $document.ReadOnly
                word_version = $word.Version
                source_sha256 = $before
                visual_review = 'Pending: inspect every rendered page; PDF export alone is not visual verification.'
            }
        } finally {
            if ($document) {
                $document.Close(0)
                [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document)
            }
        }
        if ((Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash -ne $before) { throw 'Read-only render altered the DOCX.' }
    }
} finally {
    if ($word) {
        $word.Quit(0)
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($word)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
$json = ConvertTo-Json -InputObject $results -Depth 4
$stream = [IO.File]::Open($logPath, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
try {
    $writer = [IO.StreamWriter]::new($stream, [Text.UTF8Encoding]::new($false))
    try { $writer.Write($json) } finally { $writer.Dispose() }
} finally { $stream.Dispose() }
$json
