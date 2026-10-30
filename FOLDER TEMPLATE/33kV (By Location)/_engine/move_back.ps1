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

function Get-FolderImageRange([string]$folderName) {
    $sysNum = 1
    if ($folderName -match '(?i)SISTEM\s*(\d+)|SYSTEM\s*(\d+)') {
        $sysNum = [int]($matches[1] + $matches[2])
    } elseif ($folderName -match '(?i)\b11kV\b') {
        $sysNum = 2
    } elseif ($folderName -match '(?i)\b33kV\b') {
        $sysNum = 1
    }
    $sysBase = ($sysNum - 1) * 1000

    $ancillaryMap = [ordered]@{
        'POWER TX1 - OVERVIEW' = 4001; 'POWER TX1 - COOLING FIN' = 4003; 'POWER TX1 - CONSERVATOR TANK' = 4005; 'POWER TX1 - MARSHALLING KIOSK' = 4007
        'POWER TX1 - HV CABLE COMPARTMENT' = 4009; 'POWER TX1 - HV CABLE' = 4011; 'POWER TX1 - LV CABLE COMPARTMENT' = 4013; 'POWER TX1 - LV CABLE' = 4015
        'POWER TX2 - OVERVIEW' = 4017; 'POWER TX2 - COOLING FIN' = 4019; 'POWER TX2 - CONSERVATOR TANK' = 4021; 'POWER TX2 - MARSHALLING KIOSK' = 4023
        'POWER TX2 - HV CABLE COMPARTMENT' = 4025; 'POWER TX2 - HV CABLE' = 4027; 'POWER TX2 - LV CABLE COMPARTMENT' = 4029; 'POWER TX2 - LV CABLE' = 4031
        'LOCAL TX - OVERVIEW TOP' = 4035; 'LOCAL TX - OVERVIEW' = 4033; 'LOCAL TX - HV BUSHING' = 4037; 'LOCAL TX - HV CABLE SPLIT' = 4041
        'LOCAL TX - HV CABLE' = 4039; 'LOCAL TX - LV BUSHING' = 4043; 'LOCAL TX - LV CABLE SPLIT' = 4047; 'LOCAL TX - LV CABLE' = 4045
        'FP1 - OVERVIEW' = 4049; 'FEEDER PILLAR 1' = 4049; 'FP2 - OVERVIEW' = 4051; 'FEEDER PILLAR 2' = 4051; 'FEEDER PILLAR' = 4049
        'BLACK BOX 1 - OVERVIEW' = 4053; 'BLACK BOX 2 - OVERVIEW' = 4055
        'BATTERY BANK 1 - OVERVIEW' = 4057; 'BATTERY BANK 1' = 4057; 'BATTERY BANK 2 - OVERVIEW' = 4059; 'BATTERY BANK 2' = 4059
        'BATTERY BANK 3 - OVERVIEW' = 4061; 'BATTERY BANK 3' = 4061; 'BATTERY BANK 4 - OVERVIEW' = 4063; 'BATTERY BANK 4' = 4063
        'NER - OVERVIEW' = 4065; 'N.E.R' = 4065
    }
    foreach ($k in $ancillaryMap.Keys) {
        if ($folderName -match [regex]::Escape($k)) {
            $vis = $ancillaryMap[$k]
            return @{ Visual = "$vis.jpg"; IR = "$($vis + 1).jpg" }
        }
    }

    $overviewAngles = @(
        'OVERVIEW FRONT (LHS)', 'OVERVIEW FRONT (RHS)', 'OVERVIEW REAR (LHS)', 'OVERVIEW REAR (RHS)',
        'OVERVIEW TOP (LHS)', 'OVERVIEW TOP (RHS)', 'OVERVIEW BUSBAR (LHS)', 'OVERVIEW BUSBAR (RHS)',
        'OVERVIEW BUSBAR 2 (LHS)', 'OVERVIEW BUSBAR 2 (RHS)', 'OVERVIEW CABLE CELLAR (LHS)', 'OVERVIEW CABLE CELLAR (RHS)'
    )
    for ($i = 0; $i -lt $overviewAngles.Count; $i++) {
        if ($folderName -match [regex]::Escape($overviewAngles[$i])) {
            $vis = $sysBase + 1 + ($i * 2)
            return @{ Visual = "$vis.jpg"; IR = "$($vis + 1).jpg" }
        }
    }
    if ($folderName -match '(?i)OVERVIEW') {
        $vis = $sysBase + 1
        return @{ Visual = "$vis.jpg"; IR = "$($vis + 1).jpg" }
    }

    $pNum = 1
    if ($folderName -match '(?i)PANEL\s*(\d+)') {
        $pNum = [int]$matches[1]
    }
    $compList = @(
        'BREAKER COMPARTMENT', 'BUSBAR DISCONNECTING SW', 'CABLE COMPARTMENT', 'BUSBAR COMPARTMENT',
        'PT COMPARTMENT', 'CABLE ENTRY', 'LINK BOX', 'SECONDARY COMPARTMENT'
    )
    $compIdx = 0
    for ($c = 0; $c -lt $compList.Count; $c++) {
        if ($folderName -match [regex]::Escape($compList[$c])) {
            $compIdx = $c
            break
        }
    }
    $panelBase = $sysBase + 51
    $vis = $panelBase + (($pNum - 1) * 16) + ($compIdx * 2)
    return @{ Visual = "$vis.jpg"; IR = "$($vis + 1).jpg" }
}

$subfolders = Get-ChildItem -Path $root -Directory | Where-Object { $_.Name -match '^\d+' }

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
    
    # 2. Match from folder image range (fallback)
    if (-not $targetFolder -and $f.Name -match '^\d+\.jpg$') {
        foreach ($sf in $subfolders) {
            $rng = Get-FolderImageRange $sf.Name
            if ($rng.Visual -eq $f.Name -or $rng.IR -eq $f.Name) {
                $targetFolder = $sf.FullName
                break
            }
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
