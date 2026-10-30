<#
.SYNOPSIS
    Interactive & Dynamic Folder Template Generator for Substation Type 33kV
    Supports 1, 2, 3, or 4 Switchgear Systems and up to 30 Panels per system.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [int]$SystemCount = 0,

    [Parameter(Mandatory=$false)]
    [int[]]$PanelsPerSystem,

    [Parameter(Mandatory=$false)]
    [string[]]$SystemNames,

    [Parameter(Mandatory=$false)]
    [ValidateSet('By Location', 'By Panel', 'Both')]
    [string]$Mode = '',

    [Parameter(Mandatory=$false)]
    [string]$TargetBaseDir = '',

    [Parameter(Mandatory=$false)]
    [string]$CustomTemplateName = '',

    [Parameter(Mandatory=$false)]
    [switch]$IncludeAncillary = $true
)

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$scriptRoot = $PSScriptRoot
if (-not $scriptRoot) { $scriptRoot = Get-Location }

Write-Host "===================================================" -ForegroundColor Cyan
Write-Host "   33kV SUBSTATION FOLDER TEMPLATE GENERATOR       " -ForegroundColor Yellow
Write-Host "===================================================" -ForegroundColor Cyan
Write-Host "Configurable for 1 to 4 Systems, up to 30 Panels`n" -ForegroundColor Gray

# --- 1. Interactive input if not provided via params ---

# 1.1 Number of Systems
while ($SystemCount -lt 1 -or $SystemCount -gt 4) {
    $inputSys = Read-Host "How many switchgear systems? (1, 2, 3, or 4) [Default: 1]"
    if ([string]::IsNullOrWhiteSpace($inputSys)) {
        $SystemCount = 1
    } elseif ($inputSys -match '^[1-4]$') {
        $SystemCount = [int]$inputSys
    } else {
        Write-Host "  Please enter a valid number between 1 and 4." -ForegroundColor Red
    }
}

# 1.2 System Names & Panel Counts
$systems = @()
for ($i = 1; $i -le $SystemCount; $i++) {
    $defaultName = "SISTEM $i"
    $sysName = ""
    if ($SystemNames -and ($i -le $SystemNames.Count) -and $SystemNames[$i-1]) {
        $sysName = $SystemNames[$i-1]
    } else {
        $nameInput = Read-Host "Name for System $i [Default: $defaultName]"
        $sysName = if ([string]::IsNullOrWhiteSpace($nameInput)) { $defaultName } else { $nameInput.Trim() }
    }

    $pCount = 0
    if ($PanelsPerSystem -and ($i -le $PanelsPerSystem.Count) -and $PanelsPerSystem[$i-1] -gt 0) {
        $pCount = [math]::Min(30, [math]::Max(1, $PanelsPerSystem[$i-1]))
    } else {
        while ($pCount -lt 1 -or $pCount -gt 30) {
            $pInput = Read-Host "Panel count for $sysName (1 - 30) [Default: 10]"
            if ([string]::IsNullOrWhiteSpace($pInput)) {
                $pCount = 10
            } elseif ($pInput -match '^\d+$' -and [int]$pInput -ge 1 -and [int]$pInput -le 30) {
                $pCount = [int]$pInput
            } else {
                Write-Host "  Please enter a number between 1 and 30." -ForegroundColor Red
            }
        }
    }

    $systems += [PSCustomObject]@{
        Index = $i
        Name = $sysName
        PanelCount = $pCount
    }
}

# 1.3 Mode: By Location, By Panel, or Both
if (-not $Mode) {
    Write-Host "`nSelect layout style:" -ForegroundColor White
    Write-Host "  [1] By Location (Compartment first, e.g. BREAKER - PANEL 1..N)" -ForegroundColor Gray
    Write-Host "  [2] By Panel    (Panel first, e.g. PANEL 1 - BREAKER..SECONDARY)" -ForegroundColor Gray
    Write-Host "  [3] Both        (Generate both By Location & By Panel) [Default]" -ForegroundColor Green
    $modeChoice = Read-Host "Choice (1, 2, or 3) [Default: 3]"
    switch ($modeChoice) {
        '1' { $Mode = 'By Location' }
        '2' { $Mode = 'By Panel' }
        default { $Mode = 'Both' }
    }
}

# 1.4 Target Directory
if (-not $TargetBaseDir) {
    $TargetBaseDir = $scriptRoot
}

# 1.5 Ancillary equipment confirmation
if (-not $PSBoundParameters.ContainsKey('IncludeAncillary')) {
    Write-Host "`nInclude 33kV Ancillary folders?" -ForegroundColor White
    Write-Host "  (Power TX1, Power TX2, Local TX, Feeder Pillar, Battery Bank 1-4, NER, Black Box 1-2)" -ForegroundColor Gray
    $ancInput = Read-Host "Include Ancillary equipment? (Y/N) [Default: Y]"
    if ($ancInput -match '(?i)^n') {
        $IncludeAncillary = $false
    } else {
        $IncludeAncillary = $true
    }
}

# --- 2. Definition of Compartments & Overview ---

$overviewLocations = @(
    "OVERVIEW FRONT (LHS)",
    "OVERVIEW FRONT (RHS)",
    "OVERVIEW REAR (LHS)",
    "OVERVIEW REAR (RHS)",
    "OVERVIEW TOP (LHS)",
    "OVERVIEW TOP (RHS)",
    "OVERVIEW BUSBAR (LHS)",
    "OVERVIEW BUSBAR (RHS)",
    "OVERVIEW BUSBAR (LHS)",
    "OVERVIEW BUSBAR (RHS)",
    "OVERVIEW CABLE CELLAR (LHS)",
    "OVERVIEW CABLE CELLAR (RHS)"
)

$compartments = @(
    "BREAKER COMPARTMENT",
    "BUSBAR DISCONNECTING SW",
    "CABLE COMPARTMENT",
    "BUSBAR COMPARTMENT",
    "PT COMPARTMENT",
    "CABLE ENTRY",
    "LINK BOX",
    "SECONDARY COMPARTMENT"
)

$ancillaryItems = @(
    "POWER TX1 - OVERVIEW",
    "POWER TX1 - COOLING FIN",
    "POWER TX1 - CONSERVATOR TANK",
    "POWER TX1 - MARSHALLING KIOSK",
    "POWER TX1 - HV CABLE COMPARTMENT",
    "POWER TX1 - HV CABLE",
    "POWER TX1 - LV CABLE COMPARTMENT",
    "POWER TX1 - LV CABLE",
    "POWER TX2 - OVERVIEW",
    "POWER TX2 - COOLING FIN",
    "POWER TX2 - CONSERVATOR TANK",
    "POWER TX2 - MARSHALLING KIOSK",
    "POWER TX2 - HV CABLE COMPARTMENT",
    "POWER TX2 - HV CABLE",
    "POWER TX2 - LV CABLE COMPARTMENT",
    "POWER TX2 - LV CABLE",
    "LOCAL TX - OVERVIEW",
    "LOCAL TX - OVERVIEW TOP",
    "LOCAL TX - HV BUSHING",
    "LOCAL TX - HV CABLE",
    "LOCAL TX - HV CABLE SPLIT",
    "LOCAL TX - LV BUSHING",
    "LOCAL TX - LV CABLE",
    "LOCAL TX - LV CABLE SPLIT",
    "FP1 - OVERVIEW",
    "FP2 - OVERVIEW",
    "BLACK BOX 1 - OVERVIEW",
    "BLACK BOX 2 - OVERVIEW",
    "BATTERY BANK 1 - OVERVIEW",
    "BATTERY BANK 2 - OVERVIEW",
    "BATTERY BANK 3 - OVERVIEW",
    "BATTERY BANK 4 - OVERVIEW",
    "NER - OVERVIEW"
)

# --- 3. Engine & Bat template content ---

$bat1Content = @'
@echo off
cd /d "%~dp0"
title BAT 1 - Rename Images (Master List)
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0_engine\rename_images.ps1"
echo.
pause
'@

$bat2Content = @'
@echo off
cd /d "%~dp0"
title BAT 2 - Move All Images to SORT
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0_engine\move_to_sort.ps1"
echo.
pause
'@

$bat3Content = @'
@echo off
cd /d "%~dp0"
title BAT 3 - Restore Images from SORT
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0_engine\move_back.ps1"
echo.
pause
'@

$bat4Content = @'
@echo off
cd /d "%~dp0"
title BAT 4 - Sort Images by Panel & Equipment
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0_engine\sort_by_panel.ps1"
echo.
pause
'@

$bat5Content = @'
@echo off
cd /d "%~dp0"
title BAT 5 - Move Images Back from SORT_BY_PANEL
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0_engine\move_back_from_panel.ps1"
echo.
pause
'@

$renamePs1Content = @'
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
'@


$moveToSortPs1Content = @'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
Write-Host "===================================================" -ForegroundColor Cyan
Write-Host "   BAT 2: MOVE ALL IMAGES TO SORT FOLDER           " -ForegroundColor Yellow
Write-Host "===================================================" -ForegroundColor Cyan
$root = Split-Path $PSScriptRoot -Parent
Write-Host "Working Directory: $root`n"

$sortDir = Join-Path $root "SORT"
if (-not (Test-Path $sortDir)) {
    New-Item -ItemType Directory -Path $sortDir | Out-Null
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

$movedCount = 0

foreach ($folder in ($subfolders | Sort-Object Name)) {
    $files = Get-ChildItem -Path $folder.FullName -File | Where-Object { $_.Extension -match '(?i)\.(jpg|jpeg|png)$' }
    foreach ($f in $files) {
        $destPath = Join-Path $sortDir $f.Name
        $manifest[$f.Name] = $folder.Name
        
        Move-Item -Path $f.FullName -Destination $destPath -Force
        Write-Host "  [MOVED] $($folder.Name)\$($f.Name) -> SORT\$($f.Name)" -ForegroundColor Green
        $movedCount++
    }
}

$manifest | ConvertTo-Json -Depth 2 | Set-Content -Path $manifestPath -Encoding UTF8

Write-Host "`n===================================================" -ForegroundColor Cyan
Write-Host "  COMPLETED: $movedCount images moved to:" -ForegroundColor Yellow
Write-Host "  $sortDir" -ForegroundColor White
Write-Host "===================================================" -ForegroundColor Cyan
'@

$moveBackPs1Content = @'
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
'@


$sortByPanelPs1Content = @'
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
'@


$moveBackFromPanelPs1Content = @'
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
        
        # 2. Fallback to folder image range
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
            Write-Host "  [RESTORED] SORT_BY_PANEL\$($catDir.Name)\$($f.Name) -> $destFolderName\$($f.Name)" -ForegroundColor Green
            $restoredCount++
        } else {
            Write-Host "  [SKIP] Could not determine original folder for: $($f.Name)" -ForegroundColor Yellow
        }
    }
}

$remaining = Get-ChildItem -Path $panelSortDir -Recurse -File | Where-Object { $_.Extension -match '(?i)\.(jpg|jpeg|png)$' }
if ($remaining.Count -eq 0) {
    if (Test-Path $manifestPath) { Remove-Item -Force $manifestPath }
    foreach ($catDir in $categoryDirs) {
        Remove-Item -Path $catDir.FullName -Recurse -Force
    }
}

Write-Host "`n===================================================" -ForegroundColor Cyan
Write-Host "  COMPLETED: $restoredCount images restored to their folders." -ForegroundColor Yellow
Write-Host "===================================================" -ForegroundColor Cyan
'@


# --- 4. Function to generate a template folder ---

function New-33kVTemplateFolder([string]$destPath, [string]$layoutType) {
    if (-not (Test-Path $destPath)) {
        New-Item -ItemType Directory -Path $destPath -Force | Out-Null
    }

    Write-Host "`nGenerating: $destPath ($layoutType)..." -ForegroundColor Cyan

    $folderSeq = 1
    $totalFolders = 0

    # Calculate total to determine 3-digit vs 4-digit padding
    $estFolders = 0
    foreach ($s in $systems) {
        $estFolders += $overviewLocations.Count
        $estFolders += ($s.PanelCount * $compartments.Count)
    }
    if ($IncludeAncillary) { $estFolders += $ancillaryItems.Count }

    $padFormat = if ($estFolders -ge 1000) { "{0:D4}. {1}" } else { "{0:D3}. {1}" }

    foreach ($s in $systems) {
        $sysPrefix = if ($SystemCount -gt 1) { "$($s.Name) - " } else { "" }

        # A. Overview folders
        foreach ($ov in $overviewLocations) {
            $name = $padFormat -f $folderSeq, ($sysPrefix + $ov)
            $p = Join-Path $destPath $name
            if (-not (Test-Path $p)) { New-Item -ItemType Directory -Path $p | Out-Null }
            $folderSeq++
            $totalFolders++
        }

        # B. Panels
        if ($layoutType -eq 'By Location') {
            # Compartment first, then Panel 1..N
            foreach ($comp in $compartments) {
                for ($p = 1; $p -le $s.PanelCount; $p++) {
                    $name = $padFormat -f $folderSeq, ($sysPrefix + "$comp - PANEL $p")
                    $targetP = Join-Path $destPath $name
                    if (-not (Test-Path $targetP)) { New-Item -ItemType Directory -Path $targetP | Out-Null }
                    $folderSeq++
                    $totalFolders++
                }
            }
        } else {
            # By Panel: Panel 1..N first, then Compartments
            for ($p = 1; $p -le $s.PanelCount; $p++) {
                foreach ($comp in $compartments) {
                    $name = $padFormat -f $folderSeq, ($sysPrefix + "PANEL $p - $comp")
                    $targetP = Join-Path $destPath $name
                    if (-not (Test-Path $targetP)) { New-Item -ItemType Directory -Path $targetP | Out-Null }
                    $folderSeq++
                    $totalFolders++
                }
            }
        }
    }

    # C. Ancillary Equipment
    if ($IncludeAncillary) {
        foreach ($anc in $ancillaryItems) {
            $name = $padFormat -f $folderSeq, $anc
            $targetP = Join-Path $destPath $name
            if (-not (Test-Path $targetP)) { New-Item -ItemType Directory -Path $targetP | Out-Null }
            $folderSeq++
            $totalFolders++
        }
    }

    # D. SORT folder
    $sortFolder = Join-Path $destPath "SORT"
    if (-not (Test-Path $sortFolder)) { New-Item -ItemType Directory -Path $sortFolder | Out-Null }

    # E. _engine folder & scripts
    $engineFolder = Join-Path $destPath "_engine"
    if (-not (Test-Path $engineFolder)) { New-Item -ItemType Directory -Path $engineFolder | Out-Null }

    Set-Content -Path (Join-Path $engineFolder "rename_images.ps1") -Value $renamePs1Content -Encoding UTF8
    Set-Content -Path (Join-Path $engineFolder "move_to_sort.ps1") -Value $moveToSortPs1Content -Encoding UTF8
    Set-Content -Path (Join-Path $engineFolder "move_back.ps1") -Value $moveBackPs1Content -Encoding UTF8
    Set-Content -Path (Join-Path $engineFolder "sort_by_panel.ps1") -Value $sortByPanelPs1Content -Encoding UTF8
    Set-Content -Path (Join-Path $engineFolder "move_back_from_panel.ps1") -Value $moveBackFromPanelPs1Content -Encoding UTF8

    # F. Bat files in root
    Set-Content -Path (Join-Path $destPath "1_RENAME_IMAGES.bat") -Value $bat1Content -Encoding ASCII
    Set-Content -Path (Join-Path $destPath "2_MOVE_TO_SORT.bat") -Value $bat2Content -Encoding ASCII
    Set-Content -Path (Join-Path $destPath "3_MOVE_BACK_FROM_SORT.bat") -Value $bat3Content -Encoding ASCII
    Set-Content -Path (Join-Path $destPath "4_SORT_BY_PANEL.bat") -Value $bat4Content -Encoding ASCII
    Set-Content -Path (Join-Path $destPath "5_MOVE_BACK_FROM_PANEL.bat") -Value $bat5Content -Encoding ASCII

    Write-Host "  -> Created $totalFolders numbered folders + SORT + _engine suite." -ForegroundColor Green
    return $destPath
}

# --- 5. Execute generation based on Mode ---

# Build informative folder naming
$sysDetail = if ($SystemCount -eq 1) {
    "1 Sistem ($($systems[0].PanelCount) Panels)"
} else {
    "$SystemCount Sistem (" + (($systems | ForEach-Object { "$($_.Name) - $($_.PanelCount)P" }) -join ", ") + ")"
}
$sysDetail = $sysDetail -replace '[:*?"<>|]', ''

$createdTemplates = @()

if ($Mode -eq 'By Location' -or $Mode -eq 'Both') {
    $dirName = if ($CustomTemplateName) {
        if ($Mode -eq 'Both') { "$CustomTemplateName (By Location)" } else { $CustomTemplateName }
    } else {
        "33kV $sysDetail (By Location)"
    }
    $dest = Join-Path $TargetBaseDir $dirName
    $res = New-33kVTemplateFolder -destPath $dest -layoutType 'By Location'
    $createdTemplates += $res
}

if ($Mode -eq 'By Panel' -or $Mode -eq 'Both') {
    $dirName = if ($CustomTemplateName) {
        if ($Mode -eq 'Both') { "$CustomTemplateName (By Panel)" } else { $CustomTemplateName }
    } else {
        "33kV $sysDetail (By Panel)"
    }
    $dest = Join-Path $TargetBaseDir $dirName
    $res = New-33kVTemplateFolder -destPath $dest -layoutType 'By Panel'
    $createdTemplates += $res
}

Write-Host "`n===================================================" -ForegroundColor Cyan
Write-Host "  TEMPLATE GENERATION COMPLETE!                    " -ForegroundColor Yellow
Write-Host "===================================================" -ForegroundColor Cyan
foreach ($ct in $createdTemplates) {
    Write-Host "  [CREATED] $ct" -ForegroundColor Green
}
Write-Host "===================================================" -ForegroundColor Cyan
