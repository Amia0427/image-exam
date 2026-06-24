param(
    [Parameter(Mandatory = $true)]
    [string]$RootDir
)

$ErrorActionPreference = 'Stop'

$assetsDir = Join-Path $RootDir 'assets'
$chaptersDir = Join-Path $RootDir 'chapters'
$results = @()

foreach ($assetChapter in Get-ChildItem -LiteralPath $assetsDir -Directory | Sort-Object Name) {
    $slidesDir = Join-Path $assetChapter.FullName 'slides'
    $expectedSlides = (Get-ChildItem -LiteralPath $slidesDir -File -Filter '*.png' -ErrorAction SilentlyContinue).Count
    $chapterDir = Join-Path $chaptersDir $assetChapter.Name
    $docxFiles = @(Get-ChildItem -LiteralPath $chapterDir -File -Filter '*.docx' -ErrorAction SilentlyContinue)
    $summaryPath = Join-Path $chapterDir 'chapter-summary.json'
    $summary = $null

    if (Test-Path -LiteralPath $summaryPath) {
        try {
            $summary = Get-Content -Raw -LiteralPath $summaryPath -Encoding UTF8 | ConvertFrom-Json
        } catch {
            $summary = [pscustomobject]@{
                slide_count = $null
                parse_error = $_.Exception.Message
            }
        }
    }

    $results += [pscustomobject]@{
        Chapter = $assetChapter.Name
        ExpectedSlides = $expectedSlides
        SummarySlides = if ($summary) { $summary.slide_count } else { $null }
        DocxCount = $docxFiles.Count
        HasSummary = [bool]$summary
        DocxPath = if ($docxFiles.Count -gt 0) { $docxFiles[0].FullName } else { $null }
        Status = if ($docxFiles.Count -eq 1 -and $summary -and -not $summary.parse_error -and $summary.slide_count -eq $expectedSlides) { 'OK' } else { 'CHECK' }
    }
}

$results | ConvertTo-Json -Depth 3 | Set-Content -LiteralPath (Join-Path $RootDir 'validation-summary.json') -Encoding UTF8
$results | Format-Table -AutoSize
