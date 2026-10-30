param(
    [string]$csvName = "079.csv",
    [string]$targetGroupsRaw = "all"
)

$ErrorActionPreference = "Stop"

# ── Force UTF-8 Output Encoding ───────────────────────────────────────────────
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

# ── Resolve Dynamic Paths ─────────────────────────────────────────────────────
$scriptDir = $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($scriptDir)) {
    $scriptDir = Split-Path -Path $MyInvocation.MyCommand.Path -Parent
}
$baseDir = $env:BASE_DIR
if ([string]::IsNullOrWhiteSpace($baseDir)) {
    $baseDir = Split-Path -Path $scriptDir -Parent
}

$mappingDir = $env:MAPPING_DIR
if ([string]::IsNullOrWhiteSpace($mappingDir)) { $mappingDir = Join-Path $baseDir "06. EXCELL MAPPING" }

$outputBaseDir = $env:OUTPUT_DIR
if ([string]::IsNullOrWhiteSpace($outputBaseDir)) { $outputBaseDir = Join-Path $baseDir "07. OUTPUT FILE" }

$csvStem = [System.IO.Path]::GetFileNameWithoutExtension($csvName)
$csvPath = Join-Path $mappingDir "$csvStem.csv"

if (-not (Test-Path $csvPath)) {
    Write-Error "CSV mapping file not found at: $csvPath"
    exit 1
}

# Individual pages read from FLAT INDIVIDUAL directory
$indivDir = Join-Path $outputBaseDir "INDIVIDUAL"

if (-not (Test-Path $indivDir)) {
    Write-Error "Individual pages directory not found at: $indivDir"
    exit 1
}

Write-Host "------------------------------------------------------------"
Write-Host "Starting Word Merging Pipeline..."
Write-Host "Substation Prefix : $csvStem"
Write-Host "Individual Folder : $indivDir"
Write-Host "Final Group Dir   : $outputBaseDir"
Write-Host "------------------------------------------------------------"

# ── Load CSV Mapping ──────────────────────────────────────────────────────────
$csvData = Import-Csv -Path $csvPath

# Map pages to groups based on CSV rows
# Group 0 = page0 (Executive Summary)
$groupMap = @{}
$groupMap[0] = @()

# Check Group 0 / page0 in flat INDIVIDUAL folder
$page0Path = Join-Path $indivDir "$csvStem. Group 0.docx"
if (-not (Test-Path $page0Path)) {
    $page0Path = Join-Path $indivDir "$csvStem.Group 0.docx"
}
if (-not (Test-Path $page0Path)) {
    $page0Path = Join-Path $indivDir "$csvStem. page0.docx"
}
if (-not (Test-Path $page0Path)) {
    $page0Path = Join-Path $indivDir "$csvStem.page0.docx"
}

if (Test-Path $page0Path) {
    $groupMap[0] += $page0Path
}

foreach ($row in $csvData) {
    $grpStr = $row.Group
    $outFileName = $row."Output File Name"

    if ([string]::IsNullOrWhiteSpace($grpStr) -or [string]::IsNullOrWhiteSpace($outFileName)) { continue }

    if ([int]::TryParse($grpStr.Trim(), [ref]$null)) {
        $grpNum = [int]$grpStr.Trim()
        
        if ($outFileName -match "page(\d+)") {
            $pNum = $Matches[1]
            
            # Match flat files in INDIVIDUAL directory (e.g. 079. page1.docx or 079.page1.docx)
            # Prefix with $csvStem to avoid picking up pages from OTHER substations
            $matchingFiles = Get-ChildItem -Path $indivDir -Filter "$csvStem*page$pNum.docx"
            if ($matchingFiles.Count -eq 0) {
                $matchingFiles = Get-ChildItem -Path $indivDir -Filter "$csvStem*page$pNum"
            }
            
            if ($matchingFiles.Count -gt 0) {
                if (-not $groupMap.ContainsKey($grpNum)) {
                    $groupMap[$grpNum] = @()
                }
                foreach ($f in $matchingFiles) {
                    if ($groupMap[$grpNum] -notcontains $f.FullName) {
                        $groupMap[$grpNum] += $f.FullName
                    }
                }
            }
        }
    }
}

# ── Filter Groups to Merge ────────────────────────────────────────────────────
$targetGroups = @()
if ($targetGroupsRaw -eq "all") {
    $targetGroups = $groupMap.Keys | Sort-Object
} else {
    $rawArr = $targetGroupsRaw -split '\s+|,+'
    foreach ($item in $rawArr) {
        if ([int]::TryParse($item, [ref]$null)) {
            $g = [int]$item
            if ($groupMap.ContainsKey($g)) {
                $targetGroups += $g
            }
        }
    }
}

if ($targetGroups.Count -eq 0) {
    Write-Warning "No valid groups found to merge for $csvStem."
    exit 0
}

# ── Helper: Read margins from a .docx file's XML ─────────────────────────────
function Get-DocxMargins($docxPath) {
    try {
        $zip = [System.IO.Compression.ZipFile]::OpenRead($docxPath)
        $entry = $zip.GetEntry("word/document.xml")
        $reader = New-Object System.IO.StreamReader($entry.Open())
        $xml = $reader.ReadToEnd()
        $reader.Close()
        $zip.Dispose()
        $m = [regex]::Match($xml, '<w:pgMar\s([^/]*)/?>')
        if ($m.Success) {
            $attrs = $m.Value
            $getVal = {
                param($name)
                $pattern = "w:$name=""(\d+)"""
                $r = [regex]::Match($attrs, $pattern)
                if ($r.Success) { return [int]$r.Groups[1].Value }
                return $null
            }
            return @{
                Top    = & $getVal 'top'
                Right  = & $getVal 'right'
                Bottom = & $getVal 'bottom'
                Left   = & $getVal 'left'
                Header = & $getVal 'header'
                Footer = & $getVal 'footer'
                Gutter = & $getVal 'gutter'
            }
        }
    } catch {}
    return $null
}

# -- Global Word Mutex - only ONE Word COM operation at a time ----------------
$mutexName = "Global\WRAITH_WordMerge_Mutex"
$mutex = New-Object System.Threading.Mutex($false, $mutexName)
Write-Host "[WORD LOCK] Waiting for Word mutex (another merge may be running)..."
$acquired = $mutex.WaitOne(300000)
if (-not $acquired) {
    Write-Error "[WORD LOCK] Timed out waiting for Word mutex. Another merge is still running."
    exit 1
}
Write-Host "[WORD LOCK] Mutex acquired for $csvStem - proceeding."

# ── Initialize Word COM ───────────────────────────────────────────────────────
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$word.ScreenUpdating = $false

try {
    foreach ($grp in $targetGroups) {
        $filesToMerge = $groupMap[$grp]
        if ($null -eq $filesToMerge -or $filesToMerge.Count -eq 0) { continue }

        $finalFileList = @()
        foreach ($f in $filesToMerge) {
            if ($finalFileList -notcontains $f) {
                $finalFileList += $f
            }
        }

        # Sort pages numerically
        $finalFileList = @($finalFileList | Sort-Object {
            if ($_ -match "page(\d+)\.docx") { [int]$Matches[1] } else { 0 }
        })

        # Final Merged Word File saved directly in ROOT OUTPUT FILE folder
        $outMergedPath = Join-Path $outputBaseDir "$csvStem. Group $grp.docx"
        Write-Host "Merging Group $grp ($($finalFileList.Count) pages) -> $outMergedPath"

        # Open Base Document
        $baseDoc = $word.Documents.Open($finalFileList[0])
        $selection = $word.Selection

        for ($i = 1; $i -lt $finalFileList.Count; $i++) {
            $appendFile = $finalFileList[$i]

            # Read source margins BEFORE inserting
            $srcMargins = Get-DocxMargins $appendFile

            $selection.EndKey(6) | Out-Null
            # Insert Next-Page Section Break so each appended doc gets its own section
            $selection.InsertBreak(2) | Out-Null

            # Apply the source document's margins to the new section PageSetup before inserting file
            if ($null -ne $srcMargins) {
                $pm = $selection.PageSetup
                if ($null -ne $srcMargins.Top)    { $pm.TopMargin    = [double]$srcMargins.Top / 20.0    }
                if ($null -ne $srcMargins.Bottom) { $pm.BottomMargin = [double]$srcMargins.Bottom / 20.0 }
                if ($null -ne $srcMargins.Left)   { $pm.LeftMargin   = [double]$srcMargins.Left / 20.0   }
                if ($null -ne $srcMargins.Right)  { $pm.RightMargin  = [double]$srcMargins.Right / 20.0  }
                if ($null -ne $srcMargins.Header) { $pm.HeaderDistance = [double]$srcMargins.Header / 20.0 }
                if ($null -ne $srcMargins.Footer) { $pm.FooterDistance = [double]$srcMargins.Footer / 20.0 }
                if ($null -ne $srcMargins.Gutter) { $pm.Gutter       = [double]$srcMargins.Gutter / 20.0  }
            }

            $selection.InsertFile($appendFile) | Out-Null

            try {
                $currentSection = $baseDoc.Sections.Item($baseDoc.Sections.Count)
                $currentSection.Headers.Item(1).LinkToPrevious = $false
                $currentSection.Footers.Item(1).LinkToPrevious = $false
            } catch {}
        }

        if (Test-Path $outMergedPath) { Remove-Item $outMergedPath -Force }
        $baseDoc.SaveAs([string]$outMergedPath)
        $baseDoc.Close($false)
        Write-Host " Successfully saved: $outMergedPath"
    }
} catch {
    Write-Error "Error during Word document merge: $($_.Exception.Message)"
} finally {
    # Safely release Word - it may have been killed externally
    try {
        if ($null -ne $word) {
            $word.ScreenUpdating = $true
            $word.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
        }
    } catch {
        Write-Warning "[WORD LOCK] Word COM already released or killed externally - skipping cleanup."
    } finally {
        try { $mutex.ReleaseMutex() } catch {}
        Write-Host "[WORD LOCK] Mutex released for $csvStem."
    }
}

Write-Host "Word merging pipeline finished."