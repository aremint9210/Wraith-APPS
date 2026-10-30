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
