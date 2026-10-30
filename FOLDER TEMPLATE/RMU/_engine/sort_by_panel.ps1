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

# Determine template type (RMU or VCB)
$subfolders = Get-ChildItem -Path $root -Directory | Where-Object { $_.Name -match '^\d+' }
$isRmu = ($subfolders.Count -le 55 -and ($root -match '(?i)rmu' -or -not ($subfolders.Name -match 'PANEL\s*1[1-5]')))
$maxPanels = if ($isRmu) { 10 } else { 15 }

# Master category sequence: OVERVIEW, PANEL 1..15, TX, OTHERS
$masterCategoryOrder = @("OVERVIEW")
for ($p = 1; $p -le $maxPanels; $p++) {
    $masterCategoryOrder += "PANEL $p"
}
$masterCategoryOrder += @("TX", "OTHERS")

function Get-CategoryFromContext([string]$folderName, [int]$imgNum) {
    if ($folderName -match '(?i)TX\d*|TRANSFORMER') { return 'TX' }
    if ($folderName -match '(?i)FP\d*|BLACK\s*BOX|BATTERY\s*BANK') { return 'OTHERS' }
    if ($folderName -match '(?i)PANEL\s*(\d+)') { return ('PANEL ' + [int]$matches[1]) }
    if ($folderName -match '(?i)OVERVIEW') { return 'OVERVIEW' }
    
    # Fallback using image number
    if ($imgNum -gt 0) {
        if ($isRmu) {
            if ($imgNum -le 12) { return 'OVERVIEW' }
            elseif ($imgNum -le 52) {
                $p = [math]::Floor(($imgNum - 13) / 4) + 1
                return "PANEL $p"
            }
            elseif ($imgNum -le 84) { return 'TX' }
            elseif ($imgNum -le 100) { return 'OTHERS' }
        } else {
            if ($imgNum -le 24) { return 'OVERVIEW' }
            elseif ($imgNum -le 264) {
                $p = [math]::Floor(($imgNum - 25) / 16) + 1
                return "PANEL $p"
            }
            elseif ($imgNum -le 296) { return 'TX' }
            elseif ($imgNum -le 312) { return 'OTHERS' }
        }
    }
    return 'OTHERS'
}

# 1. Discover all candidate images and determine their base categories
$itemsToMove = @()

# From numbered subfolders
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
        $origFolder = ""
        if ($sortManifest.ContainsKey($f.Name)) {
            $origFolder = $sortManifest[$f.Name]
        }
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

# 2. Determine which categories actually have images present
$presentCategories = $itemsToMove | Select-Object -ExpandProperty BaseCategory -Unique

# 3. Build ordered, numbered folder names ONLY for categories that actually have images
# Sequence: OVERVIEW, PANEL 1..15, TX, OTHERS
$categoryFolderMap = @{}
$seq = 1
foreach ($cat in $masterCategoryOrder) {
    if ($presentCategories -contains $cat) {
        $numberedName = "{0:D2}. {1}" -f $seq, $cat
        $categoryFolderMap[$cat] = $numberedName
        $seq++
    }
}

# Handle any other categories if present
foreach ($cat in $presentCategories) {
    if (-not $categoryFolderMap.ContainsKey($cat)) {
        $numberedName = "{0:D2}. {1}" -f $seq, $cat
        $categoryFolderMap[$cat] = $numberedName
        $seq++
    }
}

# 4. Clean up any existing empty directories in SORT_BY_PANEL (e.g. if previous empty folders existed)
foreach ($existingDir in (Get-ChildItem -Path $panelSortDir -Directory)) {
    $filesInDir = Get-ChildItem -Path $existingDir.FullName -Recurse -File | Where-Object { $_.Extension -match '(?i)\.(jpg|jpeg|png)$' }
    if ($filesInDir.Count -eq 0) {
        Remove-Item -Path $existingDir.FullName -Recurse -Force
    }
}

# 5. Move items into the numbered folders
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
