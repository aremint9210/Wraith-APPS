[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
Write-Host "===================================================" -ForegroundColor Cyan
Write-Host "   BAT 5: MOVE ALL IMAGES BACK FROM SORT_BY_PANEL  " -ForegroundColor Yellow
Write-Host "===================================================" -ForegroundColor Cyan
$root = Split-Path $PSScriptRoot -Parent
Write-Host "Working Directory: $root`n"

$panelSortDir = Join-Path $root "SORT_BY_PANEL"
if (-not (Test-Path $panelSortDir)) {
    Write-Host "[ERROR] SORT_BY_PANEL folder not found in $root" -ForegroundColor Red
    return
}

$manifestPath = Join-Path $panelSortDir "sort_by_panel_manifest.json"
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

# Also load reverse map if present in _engine\move_back.ps1
$reverseMap = @{}
$existingMoveBack = Join-Path $PSScriptRoot "move_back.ps1"
if (Test-Path $existingMoveBack) {
    try {
        $lines = Get-Content $existingMoveBack
        foreach ($line in $lines) {
            if ($line -match '^\s*"(\d+\.jpg)"\s*=\s*"([^"]+)"') {
                $reverseMap[$matches[1]] = $matches[2]
            }
        }
    } catch {}
}

$categoryDirs = Get-ChildItem -Path $panelSortDir -Directory
$restoredCount = 0

foreach ($catDir in $categoryDirs) {
    $files = Get-ChildItem -Path $catDir.FullName -File | Where-Object { $_.Extension -match '(?i)\.(jpg|jpeg|png)$' }
    foreach ($f in $files) {
        $targetFolder = $null
        
        # 1. Check from manifest
        if ($manifest.ContainsKey($f.Name)) {
            $entry = $manifest[$f.Name]
            $origFolderName = if ($entry.PSObject.Properties['OriginalFolder']) { $entry.OriginalFolder } else { $entry }
            $srcType = if ($entry.PSObject.Properties['SourceType']) { $entry.SourceType } else { "" }
            
            if ($origFolderName -and $origFolderName -ne "SORT") {
                $cand = Join-Path $root $origFolderName
                if (Test-Path $cand) {
                    $targetFolder = $cand
                }
            } elseif ($srcType -eq "SortFolder" -or $origFolderName -eq "SORT") {
                $sortDir = Join-Path $root "SORT"
                if (-not (Test-Path $sortDir)) { New-Item -ItemType Directory -Path $sortDir | Out-Null }
                $targetFolder = $sortDir
            }
        }
        
        # 2. Fallback to reverseMap
        if (-not $targetFolder -and $reverseMap.ContainsKey($f.Name)) {
            $cand = Join-Path $root $reverseMap[$f.Name]
            if (Test-Path $cand) { $targetFolder = $cand }
        }
        
        # 3. Fallback to Master List formula: folderIndex = ceil(M / 2)
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
            Write-Host "  [RESTORED] SORT_BY_PANEL\$($catDir.Name)\$($f.Name) -> $destFolderName\$($f.Name)" -ForegroundColor Green
            $restoredCount++
        } else {
            Write-Host "  [SKIP] Could not determine original folder for: $($f.Name)" -ForegroundColor Yellow
        }
    }
}

# Clean up SORT_BY_PANEL if empty
$remainingFiles = Get-ChildItem -Path $panelSortDir -Recurse -File | Where-Object { $_.Extension -match '(?i)\.(jpg|jpeg|png)$' }
if ($remainingFiles.Count -eq 0) {
    if (Test-Path $manifestPath) {
        Remove-Item -Force $manifestPath
    }
    foreach ($catDir in (Get-ChildItem -Path $panelSortDir -Directory)) {
        if ((Get-ChildItem -Path $catDir.FullName -Recurse).Count -eq 0) {
            Remove-Item -Path $catDir.FullName -Recurse -Force
        }
    }
    if ((Get-ChildItem -Path $panelSortDir).Count -eq 0) {
        Remove-Item -Path $panelSortDir -Recurse -Force
    }
}

Write-Host "`n===================================================" -ForegroundColor Cyan
Write-Host "  COMPLETED: $restoredCount images restored to original folders." -ForegroundColor Yellow
Write-Host "===================================================" -ForegroundColor Cyan
