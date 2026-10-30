[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
Write-Host "===================================================" -ForegroundColor Cyan
Write-Host "   BAT 1: RENAME IMAGES (33kV 1000-BLOCK STANDARD) " -ForegroundColor Yellow
Write-Host "===================================================" -ForegroundColor Cyan
$root = Split-Path $PSScriptRoot -Parent
Write-Host "Working Directory: $root`n"

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

    # 1. Ancillary equipment (starts at 4001)
    $ancillaryMap = [ordered]@{
        'POWER TX1 - OVERVIEW' = 4001
        'POWER TX1 - COOLING FIN' = 4003
        'POWER TX1 - CONSERVATOR TANK' = 4005
        'POWER TX1 - MARSHALLING KIOSK' = 4007
        'POWER TX1 - HV CABLE COMPARTMENT' = 4009
        'POWER TX1 - HV CABLE' = 4011
        'POWER TX1 - LV CABLE COMPARTMENT' = 4013
        'POWER TX1 - LV CABLE' = 4015
        'POWER TX2 - OVERVIEW' = 4017
        'POWER TX2 - COOLING FIN' = 4019
        'POWER TX2 - CONSERVATOR TANK' = 4021
        'POWER TX2 - MARSHALLING KIOSK' = 4023
        'POWER TX2 - HV CABLE COMPARTMENT' = 4025
        'POWER TX2 - HV CABLE' = 4027
        'POWER TX2 - LV CABLE COMPARTMENT' = 4029
        'POWER TX2 - LV CABLE' = 4031
        'LOCAL TX - OVERVIEW TOP' = 4035
        'LOCAL TX - OVERVIEW' = 4033
        'LOCAL TX - HV BUSHING' = 4037
        'LOCAL TX - HV CABLE SPLIT' = 4041
        'LOCAL TX - HV CABLE' = 4039
        'LOCAL TX - LV BUSHING' = 4043
        'LOCAL TX - LV CABLE SPLIT' = 4047
        'LOCAL TX - LV CABLE' = 4045
        'FP1 - OVERVIEW' = 4049
        'FEEDER PILLAR 1' = 4049
        'FP2 - OVERVIEW' = 4051
        'FEEDER PILLAR 2' = 4051
        'FEEDER PILLAR' = 4049
        'BLACK BOX 1 - OVERVIEW' = 4053
        'BLACK BOX 2 - OVERVIEW' = 4055
        'BATTERY BANK 1 - OVERVIEW' = 4057
        'BATTERY BANK 1' = 4057
        'BATTERY BANK 2 - OVERVIEW' = 4059
        'BATTERY BANK 2' = 4059
        'BATTERY BANK 3 - OVERVIEW' = 4061
        'BATTERY BANK 3' = 4061
        'BATTERY BANK 4 - OVERVIEW' = 4063
        'BATTERY BANK 4' = 4063
        'NER - OVERVIEW' = 4065
        'N.E.R' = 4065
    }

    foreach ($k in $ancillaryMap.Keys) {
        if ($folderName -match [regex]::Escape($k)) {
            $vis = $ancillaryMap[$k]
            return @{ Visual = "$vis.jpg"; IR = "$($vis + 1).jpg" }
        }
    }

    # 2. Overview views (12 standard angles)
    $overviewAngles = @(
        'OVERVIEW FRONT (LHS)', 'OVERVIEW FRONT (RHS)',
        'OVERVIEW REAR (LHS)', 'OVERVIEW REAR (RHS)',
        'OVERVIEW TOP (LHS)', 'OVERVIEW TOP (RHS)',
        'OVERVIEW BUSBAR (LHS)', 'OVERVIEW BUSBAR (RHS)',
        'OVERVIEW BUSBAR 2 (LHS)', 'OVERVIEW BUSBAR 2 (RHS)',
        'OVERVIEW CABLE CELLAR (LHS)', 'OVERVIEW CABLE CELLAR (RHS)'
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

    # 3. Panel compartments (8 standard compartments)
    $pNum = 1
    if ($folderName -match '(?i)PANEL\s*(\d+)') {
        $pNum = [int]$matches[1]
    }

    $compList = @(
        'BREAKER COMPARTMENT',
        'BUSBAR DISCONNECTING SW',
        'CABLE COMPARTMENT',
        'BUSBAR COMPARTMENT',
        'PT COMPARTMENT',
        'CABLE ENTRY',
        'LINK BOX',
        'SECONDARY COMPARTMENT'
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

if ($subfolders.Count -eq 0) {
    Write-Host "[ERROR] No numbered subfolders found in $root" -ForegroundColor Red
    return
}

$renamedCount = 0

foreach ($folder in ($subfolders | Sort-Object Name)) {
    $imgRange = Get-FolderImageRange $folder.Name
    $visualImgName = $imgRange.Visual
    $irImgName = $imgRange.IR
    
    $files = Get-ChildItem -Path $folder.FullName -File | Where-Object { $_.Extension -match '(?i)\.(jpg|jpeg|png)$' }
    if ($files.Count -eq 0) { continue }
    
    Write-Host "[$($folder.Name)] -> Expected: $visualImgName / $irImgName" -ForegroundColor Cyan
    
    $visualFile = $null
    $irFile = $null
    
    if ($files.Count -eq 1) {
        $f = $files[0]
        if ($f.Name -match '(?i)(photo|dc|vis|digital)') { $visualFile = $f } else { $irFile = $f }
    } elseif ($files.Count -ge 2) {
        foreach ($f in $files) {
            if ($f.Name -match '(?i)(photo|dc|vis|digital)') { $visualFile = $f }
            elseif ($f.Name -match '(?i)(flir|ir|thermal)') { $irFile = $f }
        }
        if (-not $visualFile -or -not $irFile) {
            $sorted = $files | Sort-Object Length -Descending
            $visualFile = $sorted[0]
            $irFile = $sorted[1]
        }
    }
    
    if ($visualFile -and $visualFile.Name -ne $visualImgName) {
        $targetPath = Join-Path $folder.FullName $visualImgName
        if (Test-Path $targetPath) { Remove-Item -Force $targetPath }
        Rename-Item -Path $visualFile.FullName -NewName $visualImgName -Force
        Write-Host "  [VISUAL] $($visualFile.Name) -> $visualImgName" -ForegroundColor Green
        $renamedCount++
    }
    
    if ($irFile -and $irFile.Name -ne $irImgName) {
        $targetPath = Join-Path $folder.FullName $irImgName
        if (Test-Path $targetPath) { Remove-Item -Force $targetPath }
        Rename-Item -Path $irFile.FullName -NewName $irImgName -Force
        Write-Host "  [IR]     $($irFile.Name) -> $irImgName" -ForegroundColor Green
        $renamedCount++
    }
}

Write-Host "`n===================================================" -ForegroundColor Cyan
Write-Host "  COMPLETED: $renamedCount images renamed in their folders." -ForegroundColor Yellow
Write-Host "===================================================" -ForegroundColor Cyan
