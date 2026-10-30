[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
Write-Host "===================================================" -ForegroundColor Cyan
Write-Host "   BAT 3: MOVE ALL IMAGES FROM SORT FOLDER BACK    " -ForegroundColor Yellow
Write-Host "===================================================" -ForegroundColor Cyan
$root = Split-Path $PSScriptRoot -Parent
Write-Host "Working Directory: $root`n"

$sortDir = Join-Path $root "SORT"
if (-not (Test-Path $sortDir)) {
    Write-Host "[ERROR] SORT folder not found in $root" -ForegroundColor Red
    return
}

$manifestPath = Join-Path $sortDir "sort_manifest.json"
$manifest = @{}
if (Test-Path $manifestPath) {
    try {
        $jsonContent = Get-Content -Path $manifestPath -Raw | ConvertFrom-Json
        foreach ($prop in $jsonContent.PSObject.Properties) {
            $manifest[$prop.Name] = $prop.Value
        }
    } catch {}
}

$subfolders = Get-ChildItem -Path $root -Directory | Where-Object { $_.Name -match '^\d+' }
$folderMap = @{}
foreach ($sf in $subfolders) {
    if ($sf.Name -match '^(\d+)') {
        $idx = [int]$matches[1]
        $folderMap[$idx] = $sf
    }
}

$filesInSort = Get-ChildItem -Path $sortDir -File | Where-Object { $_.Extension -match '(?i)\.(jpg|jpeg|png)$' }

if ($filesInSort.Count -eq 0) {
    Write-Host "[INFO] No image files found in SORT folder." -ForegroundColor Yellow
    return
}

$restoredCount = 0

foreach ($f in $filesInSort) {
    $targetFolder = $null
    
    # 1. Match from manifest
    if ($manifest.ContainsKey($f.Name)) {
        $targetFolderName = $manifest[$f.Name]
        $cand = Join-Path $root $targetFolderName
        if (Test-Path $cand) {
            $targetFolder = $cand
        }
    }
    
    # 2. Match from Master List formula: folder index = ceil(M / 2)
    if (-not $targetFolder -and $f.Name -match '^(\d+)\.jpg$') {
        $imgNum = [int]$matches[1]
        $folderIndex = [math]::Ceiling($imgNum / 2)
        if ($folderMap.ContainsKey($folderIndex)) {
            $targetFolder = $folderMap[$folderIndex].FullName
        }
    }
    
    if ($targetFolder) {
        $destPath = Join-Path $targetFolder $f.Name
        Move-Item -Path $f.FullName -Destination $destPath -Force
        $destFolderName = Split-Path $targetFolder -Leaf
        Write-Host "  [RESTORED] SORT\$($f.Name) -> $destFolderName\$($f.Name)" -ForegroundColor Green
        $restoredCount++
    } else {
        Write-Host "  [SKIP] Could not determine original folder for: $($f.Name)" -ForegroundColor Yellow
    }
}

$remaining = Get-ChildItem -Path $sortDir -File | Where-Object { $_.Extension -match '(?i)\.(jpg|jpeg|png)$' }
if ($remaining.Count -eq 0 -and (Test-Path $manifestPath)) {
    Remove-Item -Force $manifestPath
}

Write-Host "`n===================================================" -ForegroundColor Cyan
Write-Host "  COMPLETED: $restoredCount images restored to their folders." -ForegroundColor Yellow
Write-Host "===================================================" -ForegroundColor Cyan
