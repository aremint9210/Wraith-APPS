param(
    [string]$choices = "all",
    [string]$pages = "all",
    [string]$baseDir = "",
    [string]$SubstationDir = ""
)

if ([string]::IsNullOrWhiteSpace($baseDir) -and -not [string]::IsNullOrWhiteSpace($SubstationDir)) {
    $baseDir = $SubstationDir
}

$scriptDir = $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($scriptDir)) {
    $scriptDir = Split-Path -Path $MyInvocation.MyCommand.Path -Parent
}

if (-not [string]::IsNullOrWhiteSpace($baseDir)) {
    $env:BASE_DIR = $baseDir
}

if ([string]::IsNullOrWhiteSpace($env:BASE_DIR)) {
    $targetSubstationDir = Split-Path -Path $scriptDir -Parent
} else {
    $targetSubstationDir = $env:BASE_DIR
}

# Resolve MAPPING_DIR
$MAPPING_DIR = Join-Path $targetSubstationDir "06. EXCELL MAPPING"
if (-not (Test-Path $MAPPING_DIR)) {
    $MAPPING_DIR = Join-Path $targetSubstationDir "EXCELL MAPPING"
}

if (-not (Test-Path $MAPPING_DIR)) {
    Write-Host "❌ Error: Cannot find '06. EXCELL MAPPING' directory in: $targetSubstationDir" -ForegroundColor Red
    return
}

$generateScript = Join-Path $scriptDir "generate_report.py"
$mergeScript = Join-Path $scriptDir "dynamic_merge_tool.ps1"

# 1. Get List of CSV Files
$csvFiles = Get-ChildItem -Path $MAPPING_DIR -Filter "*.csv" | Where-Object { 
    $_.Name -notin "image number.csv", "Master List.csv" 
} | Sort-Object Name

# 2. Determine which files to process
$targetFiles = @()
if ($choices -eq "all") {
    $targetFiles = $csvFiles
} else {
    $indices = [regex]::Split($choices, '[\s,;]+') | Where-Object { $_ -ne "" } | ForEach-Object { [int]$_ - 1 }
    foreach ($idx in $indices) {
        if ($idx -ge 0 -and $idx -lt $csvFiles.Count) {
            $targetFiles += $csvFiles[$idx]
        }
    }
}

if ($targetFiles.Count -eq 0) {
    Write-Host "No valid CSV files selected." -ForegroundColor Red
    return
}

Write-Host "`nFound $($targetFiles.Count) substations to process." -ForegroundColor Yellow

# 3. Process each file one-by-one
foreach ($file in $targetFiles) {
    $csvName = $file.Name
    
    Write-Host "`n===============================================" -ForegroundColor Yellow
    Write-Host "   NOW PROCESSING: $csvName" -ForegroundColor Yellow
    Write-Host "===============================================" -ForegroundColor Yellow

    # STEP 1: GENERATION
    Write-Host "`n>>> STEP 1: GENERATING REPORTS ($csvName) <<<" -ForegroundColor Cyan
    & python $generateScript $csvName $pages "--no-merge"
    if ($LASTEXITCODE -ne 0) {
        Write-Host "Warning: Generator reported errors for $csvName" -ForegroundColor Yellow
    }

    # STEP 2: MERGING
    Write-Host "`n>>> STEP 2: MERGING GROUPS ($csvName) <<<" -ForegroundColor Cyan
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $mergeScript -csvName $csvName -targetGroupsRaw "all"
    if ($LASTEXITCODE -ne 0) {
        Write-Host "Warning: Merger reported errors for $csvName" -ForegroundColor Yellow
    }

    Write-Host "`n[COMPLETED] Substation: $csvName" -ForegroundColor Green
    Write-Host "-----------------------------------------------"
}

Write-Host "`n===============================================" -ForegroundColor Green
Write-Host "   ALL SUBSTATIONS PROCESSED SUCCESSFULLY" -ForegroundColor Green
Write-Host "===============================================" -ForegroundColor Green
