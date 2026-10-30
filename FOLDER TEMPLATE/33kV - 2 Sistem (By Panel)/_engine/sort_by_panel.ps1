[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
Write-Host "===================================================" -ForegroundColor Cyan
Write-Host "   BAT 4: SORT IMAGES BY PANEL & EQUIPMENT         " -ForegroundColor Yellow
Write-Host "===================================================" -ForegroundColor Cyan
$root = Split-Path $PSScriptRoot -Parent
Write-Host "Working Directory: $root`n"

$panelSortDir = Join-Path $root "SORT_BY_PANEL"
if (-not (Test-Path $panelSortDir)) {
    New-Item -ItemType Directory -Path $panelSortDir | Out-Null
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

# Dynamic Category Extractor supporting 1000-block standard (S1..S4, Panels 1..30, Ancillary 4001+)
function Get-CategoryFromContext([string]$folderName, [int]$imgNum) {
    if ($folderName -match '(?i)POWER\s*TX\s*(\d+)') { return "POWER TX $($matches[1])" }
    if ($folderName -match '(?i)LOCAL\s*TX|LTX') { return "LOCAL TX" }
    if ($folderName -match '(?i)\bTX\d*|TRANSFORMER') { return "TX" }
    if ($folderName -match '(?i)FEEDER\s*PILLAR|\bFP\d*') { return "FEEDER PILLAR" }
    if ($folderName -match '(?i)BATTERY\s*BANK\s*(\d+)') { return "BATTERY BANK $($matches[1])" }
    if ($folderName -match '(?i)\bNER\b|N\.E\.R') { return "NER" }
    if ($folderName -match '(?i)BLACK\s*BOX') { return "BLACK BOX" }

    $sysPrefix = ""
    if ($folderName -match '(?i)(?:SISTEM|SYSTEM)\s*(\d+)') {
        $sysPrefix = "SISTEM $($matches[1]) - "
    } elseif ($folderName -match '(?i)(?:33kV|11kV|132kV|22kV)') {
        $sysPrefix = "$($matches[0].ToUpper()) - "
    }

    if ($folderName -match '(?i)PANEL\s*(\d+)') {
        return ($sysPrefix + "PANEL " + [int]$matches[1])
    }

    if ($folderName -match '(?i)OVERVIEW') {
        return ($sysPrefix + "OVERVIEW")
    }

    # If folder name is unknown or empty, infer from 1000-block image number
    if ($imgNum -gt 0) {
        if ($imgNum -ge 4001) {
            if ($imgNum -le 4016) { return "POWER TX 1" }
            if ($imgNum -le 4032) { return "POWER TX 2" }
            if ($imgNum -le 4048) { return "LOCAL TX" }
            if ($imgNum -le 4052) { return "FEEDER PILLAR" }
            if ($imgNum -le 4060) { return "BATTERY BANK" }
            if ($imgNum -le 4062) { return "NER" }
            if ($imgNum -le 4066) { return "BLACK BOX" }
            return "OTHERS"
        }
        
        $sIdx = [math]::Floor(($imgNum - 1) / 1000) + 1
        $sBase = ($sIdx - 1) * 1000
        $rem = $imgNum - $sBase
        $pPrefix = "SISTEM $sIdx - "
        if ($rem -le 50) {
            return ($pPrefix + "OVERVIEW")
        } else {
            $pNum = [math]::Floor(($rem - 51) / 16) + 1
            return ($pPrefix + "PANEL $pNum")
        }
    }

    return "OTHERS"
}

# 1. Discover all candidate images
$itemsToMove = @()

foreach ($folder in ($subfolders | Sort-Object Name)) {
    $files = Get-ChildItem -Path $folder.FullName -File | Where-Object { $_.Extension -match '(?i)\.(jpg|jpeg|png)$' }
    foreach ($f in $files) {
        $imgNum = 0
        if ($f.BaseName -match '^(\d+)$') { $imgNum = [int]$matches[1] }
        $baseCat = Get-CategoryFromContext $folder.Name $imgNum
        $itemsToMove += [PSCustomObject]@{
            File = $f
            BaseCategory = $baseCat
            OriginalFolder = $folder.Name
            SourceType = "Subfolder"
        }
    }
}

# From SORT folder
$sortDir = Join-Path $root "SORT"
if (Test-Path $sortDir) {
    $sortManifestPath = Join-Path $sortDir "sort_manifest.json"
    $sortManifest = @{}
    if (Test-Path $sortManifestPath) {
        try {
            $json = Get-Content -Path $sortManifestPath -Raw | ConvertFrom-Json
            foreach ($prop in $json.PSObject.Properties) {
                $sortManifest[$prop.Name] = $prop.Value
            }
        } catch {}
    }
    
    $sortFiles = Get-ChildItem -Path $sortDir -File | Where-Object { $_.Extension -match '(?i)\.(jpg|jpeg|png)$' }
    foreach ($f in $sortFiles) {
        $origFolder = if ($sortManifest.ContainsKey($f.Name)) { $sortManifest[$f.Name] } else { "" }
        $imgNum = 0
        if ($f.BaseName -match '^(\d+)$') { $imgNum = [int]$matches[1] }
        $baseCat = Get-CategoryFromContext $origFolder $imgNum
        $itemsToMove += [PSCustomObject]@{
            File = $f
            BaseCategory = $baseCat
            OriginalFolder = $origFolder
            SourceType = "SortFolder"
        }
    }
}

if ($itemsToMove.Count -eq 0) {
    Write-Host "[INFO] No image files found to sort in $root or SORT folder." -ForegroundColor Yellow
    return
}

# 2. Determine present categories and build master order
$presentCategories = $itemsToMove | Select-Object -ExpandProperty BaseCategory -Unique

# Build ordered master sequence (Sistem 1..4 with Panel 1..30, then TX, OTHERS)
$masterCategoryOrder = @()
for ($s = 1; $s -le 4; $s++) {
    $masterCategoryOrder += "SISTEM $s - OVERVIEW"
    for ($p = 1; $p -le 30; $p++) {
        $masterCategoryOrder += "SISTEM $s - PANEL $p"
    }
}
$masterCategoryOrder += "OVERVIEW"
for ($p = 1; $p -le 30; $p++) {
    $masterCategoryOrder += "PANEL $p"
}
$masterCategoryOrder += @(
    "POWER TX 1", "POWER TX 2", "POWER TX 3", "POWER TX 4",
    "LOCAL TX", "TX",
    "FEEDER PILLAR",
    "BATTERY BANK 1", "BATTERY BANK 2", "BATTERY BANK 3", "BATTERY BANK 4", "BATTERY BANK",
    "NER", "BLACK BOX", "OTHERS"
)

$categoryFolderMap = @{}
$seq = 1
foreach ($cat in $masterCategoryOrder) {
    if ($presentCategories -contains $cat) {
        $numberedName = "{0:D2}. {1}" -f $seq, $cat
        $categoryFolderMap[$cat] = $numberedName
        $seq++
    }
}
foreach ($cat in $presentCategories) {
    if (-not $categoryFolderMap.ContainsKey($cat)) {
        $numberedName = "{0:D2}. {1}" -f $seq, $cat
        $categoryFolderMap[$cat] = $numberedName
        $seq++
    }
}

# 3. Clean up existing empty directories
foreach ($existingDir in (Get-ChildItem -Path $panelSortDir -Directory)) {
    $filesInDir = Get-ChildItem -Path $existingDir.FullName -Recurse -File | Where-Object { $_.Extension -match '(?i)\.(jpg|jpeg|png)$' }
    if ($filesInDir.Count -eq 0) {
        Remove-Item -Path $existingDir.FullName -Recurse -Force
    }
}

# 4. Move items into numbered folders
$movedCount = 0
$stats = @{}

foreach ($item in $itemsToMove) {
    $f = $item.File
    $cat = $item.BaseCategory
    $folderName = $categoryFolderMap[$cat]
    $targetDir = Join-Path $panelSortDir $folderName
    if (-not (Test-Path $targetDir)) {
        New-Item -ItemType Directory -Path $targetDir | Out-Null
    }
    
    $destPath = Join-Path $targetDir $f.Name
    $manifest[$f.Name] = @{
        BaseCategory = $cat
        FolderName = $folderName
        OriginalFolder = $item.OriginalFolder
        SourceType = $item.SourceType
    }
    
    Move-Item -Path $f.FullName -Destination $destPath -Force
    Write-Host "  [MOVED] $($f.Name) -> SORT_BY_PANEL\$folderName\$($f.Name)" -ForegroundColor Green
    $movedCount++
    $stats[$folderName] = ($stats[$folderName] + 1)
}

$manifest | ConvertTo-Json -Depth 3 | Set-Content -Path $manifestPath -Encoding UTF8

Write-Host "`n===================================================" -ForegroundColor Cyan
Write-Host "  COMPLETED: $movedCount images organized into SORT_BY_PANEL" -ForegroundColor Yellow
Write-Host "===================================================" -ForegroundColor Cyan
foreach ($key in ($stats.Keys | Sort-Object)) {
    Write-Host "  $key : $($stats[$key]) images" -ForegroundColor White
}
Write-Host "===================================================" -ForegroundColor Cyan
