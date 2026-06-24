param(
    [Parameter(Mandatory = $true)]
    [string]$InputDir,

    [Parameter(Mandatory = $true)]
    [string]$OutputDir
)

$ErrorActionPreference = 'Stop'

$msoTrue = -1
$msoFalse = 0
$msoPicture = 13
$msoLinkedPicture = 11
$ppShapeFormatPNG = 2
$ppSaveAsPNG = 18
$tempRoot = 'C:\codex_ppt_export'

function Get-SafeName {
    param([string]$Name)
    $safe = [System.IO.Path]::GetFileNameWithoutExtension($Name)
    foreach ($char in [System.IO.Path]::GetInvalidFileNameChars()) {
        $safe = $safe.Replace($char, '_')
    }
    return $safe
}

New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
New-Item -ItemType Directory -Force -Path $tempRoot | Out-Null

$pptFiles = Get-ChildItem -LiteralPath $InputDir -File |
    Where-Object { $_.Extension -in '.ppt', '.pptx', '.pps', '.ppsx' } |
    Sort-Object Name

$powerPoint = New-Object -ComObject PowerPoint.Application
$powerPoint.Visible = $msoTrue

$summary = @()

try {
    foreach ($ppt in $pptFiles) {
        $chapterName = Get-SafeName -Name $ppt.Name
        $chapterDir = Join-Path $OutputDir $chapterName
        $slidesDir = Join-Path $chapterDir 'slides'
        $imagesDir = Join-Path $chapterDir 'images'

        New-Item -ItemType Directory -Force -Path $slidesDir | Out-Null
        New-Item -ItemType Directory -Force -Path $imagesDir | Out-Null

        Write-Host "Processing $($ppt.Name)"
        $presentation = $null
        try {
            $presentation = $powerPoint.Presentations.Open($ppt.FullName, $msoTrue, $msoFalse, $msoTrue)
        } catch {
            throw "Failed to open $($ppt.FullName): $($_.Exception.Message)"
        }
        try {
            $rawSlidesDir = Join-Path $tempRoot ([guid]::NewGuid().ToString('N'))
            New-Item -ItemType Directory -Force -Path $rawSlidesDir | Out-Null
            $presentation.SaveAs($rawSlidesDir, $ppSaveAsPNG)

            $rawSlideFiles = Get-ChildItem -LiteralPath $rawSlidesDir -File -Filter '*.png' |
                Sort-Object {
                    if ($_.BaseName -match '(\d+)$') { [int]$Matches[1] } else { [int]::MaxValue }
                }, Name

            $slideNumber = 0
            foreach ($rawSlideFile in $rawSlideFiles) {
                $slideNumber += 1
                Copy-Item -LiteralPath $rawSlideFile.FullName -Destination (Join-Path $slidesDir ('slide-{0:D3}.png' -f $slideNumber)) -Force
            }

            foreach ($slide in $presentation.Slides) {
                $shapeIndex = 0
                foreach ($shape in $slide.Shapes) {
                    $shapeIndex += 1
                    if ($shape.Type -eq $msoPicture -or
                        $shape.Type -eq $msoLinkedPicture) {
                        $imageFile = Join-Path $imagesDir ('slide-{0:D3}-shape-{1:D2}.png' -f $slide.SlideIndex, $shapeIndex)
                        try {
                            $shape.Export($imageFile, $ppShapeFormatPNG)
                        } catch {
                            $imageFile = $null
                        }
                    }
                }
            }

            $summary += [pscustomobject]@{
                File = $ppt.Name
                Chapter = $chapterName
                Slides = $presentation.Slides.Count
                SlideImages = (Get-ChildItem -LiteralPath $slidesDir -File -Filter '*.png').Count
                ExtractedImages = (Get-ChildItem -LiteralPath $imagesDir -File -Filter '*.png').Count
            }
        } finally {
            if ($presentation) {
                try {
                    $presentation.Close()
                } catch {
                    Write-Warning "PowerPoint reported a close warning for $($ppt.Name): $($_.Exception.Message)"
                }
            }
        }
    }
} finally {
    $powerPoint.Quit()
}

$summary | ConvertTo-Json -Depth 3 | Set-Content -LiteralPath (Join-Path $OutputDir 'asset-summary.json') -Encoding UTF8
$summary | Format-Table -AutoSize
