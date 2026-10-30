[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
Write-Host "===================================================" -ForegroundColor Cyan
Write-Host "   BAT 1: RENAME IMAGES IN RESPECTIVE FOLDERS      " -ForegroundColor Yellow
Write-Host "===================================================" -ForegroundColor Cyan
$root = Split-Path $PSScriptRoot -Parent
Write-Host "Working Directory: $root`n"

$subfolders = Get-ChildItem -Path $root -Directory | Where-Object { $_.Name -match '^\d+' }

if ($subfolders.Count -eq 0) {
    Write-Host "[ERROR] No numbered subfolders found in $root" -ForegroundColor Red
    return
}

$renamedCount = 0

foreach ($folder in ($subfolders | Sort-Object Name)) {
    if ($folder.Name -match '^(\d+)') {
        $no = [int]$matches[1]
        $visualImgName = "$((($no * 2) - 1)).jpg"
        $irImgName = "$($no * 2).jpg"
        
        $files = Get-ChildItem -Path $folder.FullName -File | Where-Object { $_.Extension -match '(?i)\.(jpg|jpeg|png)$' }
        if ($files.Count -eq 0) { continue }
        
        Write-Host "[$($folder.Name)]" -ForegroundColor Cyan
        
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
}

Write-Host "`n===================================================" -ForegroundColor Cyan
Write-Host "  COMPLETED: $renamedCount images renamed in their folders." -ForegroundColor Yellow
Write-Host "===================================================" -ForegroundColor Cyan
