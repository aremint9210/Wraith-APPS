param(
    [Parameter(Mandatory = $true)]
    [string]$ManifestPath,
    [string]$SubstationFilter = ""
)

$ErrorActionPreference = "Stop"

function Resolve-FlirBinPath {
    $scriptDir = $PSScriptRoot
    if ([string]::IsNullOrWhiteSpace($scriptDir)) {
        $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
    }
    if ($scriptDir -and (Test-Path (Join-Path $scriptDir "flir_libs"))) {
        return (Join-Path $scriptDir "flir_libs")
    }

    $registryPath = "Registry::HKEY_CLASSES_ROOT\WOW6432Node\CLSID\{0A235AB1-0B24-4808-BC7F-B631374B5E13}\InprocServer32"
    $defaultBinPath = "C:\Program Files (x86)\FLIR Systems\FLIR Word Reporting\bin"
    try {
        $registeredDll = (Get-ItemProperty -Path $registryPath -ErrorAction Stop)."(default)"
        if ($registeredDll -and (Test-Path $registeredDll)) {
            return Split-Path -Path $registeredDll -Parent
        }
    } catch {}
    if (Test-Path $defaultBinPath) { return $defaultBinPath }
    throw "FLIR Word Reporting installation was not found."
}

function Apply-FieldOverrides {
    param($Document, $Manifest)
    $overrides = @{}
    foreach ($page in $Manifest.pages) {
        $properties = $page.field_overrides.PSObject.Properties
        if ($null -ne $properties) {
            foreach ($property in $properties) {
                $overrides[$property.Name] = [string]$property.Value      
            }
        }
    }
    $selection = $Document.Content
    foreach ($key in $overrides.Keys) {
        if ($key -eq "{{visual_image}}") { continue }
        $val = $overrides[$key]
        if ($key -eq "{{ panel.heateramp }}" -and $overrides.ContainsKey("{{ swg.type }}") -and $overrides["{{ swg.type }}"] -match "VCB") {
            $range = $Document.Content
            $range.Find.ClearFormatting()
            while ($range.Find.Execute($key, $false, $true, $false, $false, $false, $true, 1, $false, "", 0)) {
                $currentSize = $range.Font.Size
                $range.Text = $val
                if ($currentSize -gt 1 -and $currentSize -lt 100) {
                    $range.Font.Size = $currentSize - 1
                }
                $range.Collapse(0)
            }
            continue
        }
        $selection.Find.ClearFormatting()
        $selection.Find.Replacement.ClearFormatting()
        $hasFormat = $false
        if ($val -and $val.Trim().ToUpper() -eq "HERMETICALLY SEAL") {
            $selection.Find.Replacement.Font.Size = 7
            $hasFormat = $true
        }
        $selection.Find.Execute($key, $false, $true, $false, $false, $false, $true, 1, $hasFormat, $val, 2) | Out-Null
    }
}

function Replace-TextWithImage {
    param($Document, $Manifest)
    foreach ($page in $Manifest.pages) {
        if ([string]::IsNullOrWhiteSpace($page.visual_path) -or -not (Test-Path $page.visual_path)) { continue }
        $range = $Document.Content
        $range.Find.ClearFormatting()
        if ($range.Find.Execute("{{visual_image}}")) {
            $imageRange = $range.Duplicate
            $imageRange.ParagraphFormat.Alignment = 1 # Center
            $range.Text = "" 
            $shape = $imageRange.InlineShapes.AddPicture($page.visual_path, $false, $true)
            if ($shape.Width -gt 400) {
                $ratio = 400 / $shape.Width
                $shape.Width = 400
                $shape.Height = $shape.Height * $ratio
            }
        }
    }
}

function Remove-EmptyParagraphsInHeadersAndFooters {
    param($Document)

    foreach ($section in $Document.Sections) {
        try {
            $section.PageSetup.HeaderDistance = 0
            $section.PageSetup.FooterDistance = 0
        } catch {}

        for ($h = 1; $h -le 3; $h++) {
            try { $section.Headers.Item($h).Range.Delete() } catch {}
            try { $section.Footers.Item($h).Range.Delete() } catch {}
            
            try {
                $section.Headers.Item($h).Range.Font.Size = 1
                $section.Headers.Item($h).Range.Font.Hidden = $true
                $section.Headers.Item($h).Range.ParagraphFormat.SpaceBefore = 0
                $section.Headers.Item($h).Range.ParagraphFormat.SpaceAfter = 0
                $section.Headers.Item($h).Range.ParagraphFormat.LineSpacingRule = 4
                $section.Headers.Item($h).Range.ParagraphFormat.LineSpacing = 1
            } catch {}
            
            try {
                $section.Footers.Item($h).Range.Font.Size = 1
                $section.Footers.Item($h).Range.ParagraphFormat.SpaceBefore = 0
                $section.Footers.Item($h).Range.ParagraphFormat.SpaceAfter = 0
                $section.Footers.Item($h).Range.ParagraphFormat.LineSpacingRule = 4
                $section.Footers.Item($h).Range.ParagraphFormat.LineSpacing = 1
            } catch {}
        }
    }
}

function Remove-EmptyParagraphsAtEdges {
    param($Document)
    $paragraphs = $Document.Paragraphs
    
    # Remove from end (bottom of document)
    for ($i = 0; $i -lt 10; $i++) {
        if ($paragraphs.Count -le 1) { break }
        $lastPara = $paragraphs.Item($paragraphs.Count)
        if ($lastPara.Range.Tables.Count -eq 0) {
            $t = $lastPara.Range.Text
            if ($null -ne $t -and $t.Replace("`r","").Replace("`n","").Replace("`a","").Trim().Length -eq 0) {
                try { $lastPara.Range.Delete() } catch { break }
            } else { break }
        } else { break }
    }

    # Remove from start (top of document)
    for ($i = 0; $i -lt 10; $i++) {
        if ($paragraphs.Count -le 1) { break }
        $firstPara = $paragraphs.Item(1)
        if ($firstPara.Range.Tables.Count -eq 0) {
            $t = $firstPara.Range.Text
            if ($null -ne $t -and $t.Replace("`r","").Replace("`n","").Replace("`a","").Trim().Length -eq 0) {
                try { $firstPara.Range.Delete() } catch { break }
            } else { break }
        } else { break }
    }
}

function Remove-PageBreaksAndMerge {
    param($Document)
    $range = $Document.Content
    $range.Find.ClearFormatting()
    $range.Find.Replacement.ClearFormatting()
    $range.Find.Execute("^m", $false, $false, $false, $false, $false, $true, 1, $false, "", 2) | Out-Null
    $range.Find.Execute("^b", $false, $false, $false, $false, $false, $true, 1, $false, "", 2) | Out-Null
    
    $paragraphs = $Document.Paragraphs
    $deleteIndices = @()
    $isBelowTable = $false
    
    for ($i = $paragraphs.Count; $i -ge 1; $i--) {
        $currPara = $paragraphs.Item($i)
        if ($currPara.Range.Tables.Count -gt 0) {
            $isBelowTable = $true
        } else {
            $text = $currPara.Range.Text
            if ($null -ne $text) {
                $cleanText = $text.Replace("`r","").Replace("`n","").Replace("`a","").Trim()
                if ($cleanText.Length -eq 0) {
                    if ($isBelowTable) {
                        $hasTableAbove = $false
                        for ($j = $i - 1; $j -ge 1; $j--) {
                            $checkPara = $paragraphs.Item($j)
                            if ($checkPara.Range.Tables.Count -gt 0) {
                                $hasTableAbove = $true
                                break
                            }
                            $cText = $checkPara.Range.Text
                            if ($null -ne $cText) {
                                $clText = $cText.Replace("`r","").Replace("`n","").Replace("`a","").Trim()
                                if ($clText.Length -gt 0) { break }
                            }
                        }
                        if ($hasTableAbove) {
                            $deleteIndices += $i
                        }
                    }
                } else {
                    $isBelowTable = $false
                }
            }
        }
    }
    
    foreach ($idx in $deleteIndices) {
        try { $paragraphs.Item($idx).Range.Delete() } catch {}
    }
}

function New-ProReport {
    param($Manifest, $Word)
    $totalTemplatePages = 1
    try {
        $tempDoc = $Word.Documents.Open($Manifest.template_path, $false, $true)
        $totalTemplatePages = $tempDoc.ComputeStatistics(2)
        $tempDoc.Close($false)
    } catch {
        Write-Warning "Could not count template pages. Defaulting to 1."
    }
    $report = New-Object WizardClassLibrary.ProReport
    $report.Template = [string]$Manifest.template_path
    $report.IntroPageNumbers = New-Object "System.Collections.Generic.List[System.Int32]"
    $report.SummaryPageNumbers = New-Object "System.Collections.Generic.List[System.Int32]"
    $report.ReportPages = New-Object "System.Collections.Generic.List[WizardClassLibrary.ProReportPage]"
    $report.ReportProperties = New-Object "System.Collections.Generic.List[WizardClassLibrary.ReportProperty]"
    for ($i = 1; $i -le $totalTemplatePages; $i++) {
        $reportPage = New-Object WizardClassLibrary.ProReportPage
        $reportPage.TemplatePageNr = $i
        $reportPage.IRImages = New-Object "System.Collections.Generic.List[WizardClassLibrary.ProIRImage]"
        $reportPage.DCImages = New-Object "System.Collections.Generic.List[WizardClassLibrary.ProDCImage]"
        if ($i -eq 1) {
            $firstPage = $Manifest.pages[0]
            if (-not [string]::IsNullOrWhiteSpace($firstPage.ir_path)) {
                $irImage = New-Object WizardClassLibrary.ProIRImage
                $irImage.FileName = [string]$firstPage.ir_path
                $irImage.Sketch = 1
                $irImage.TextComments = New-Object "System.Collections.Generic.List[WizardClassLibrary.TextComment]"
                $reportPage.IRImages.Add($irImage)
            }
            if (-not [string]::IsNullOrWhiteSpace($firstPage.visual_path)) {
                $dcImage = New-Object WizardClassLibrary.ProDCImage
                $dcImage.FileName = [string]$firstPage.visual_path
                $dcImage.Sketch = 0
                $reportPage.DCImages.Add($dcImage)
            }
        }
        $report.ReportPages.Add($reportPage)
    }
    return $report
}

function Serialize-ProReport {
    param($Report, [string]$OutputPath)
    $serializer = New-Object System.Xml.Serialization.XmlSerializer([WizardClassLibrary.ProReport])
    $writer = New-Object System.IO.StreamWriter($OutputPath, $false, [System.Text.Encoding]::Unicode)
    try { $serializer.Serialize($writer, $Report) } finally { $writer.Dispose() }
}

$document = $null
$word = $null
try {
    $manifestArray = Get-Content -Raw -Path $ManifestPath | ConvertFrom-Json   
    $flirBinPath = Resolve-FlirBinPath
    $wizardDll = Join-Path $flirBinPath "WizardClassLibrary.dll"
    $wordAddinDll = Join-Path $flirBinPath "WordAddin_O2k7.dll"
    [Reflection.Assembly]::LoadFrom($wizardDll) | Out-Null
    # Ensure FLIR Add-in is registered & enabled in HKCU registry before starting Word
    $regPath = "HKCU:\Software\Microsoft\Office\Word\Addins\WordAddin_O2k7.WordAddinEngine"
    if (-not (Test-Path $regPath)) { New-Item -Path $regPath -Force | Out-Null }
    Set-ItemProperty -Path $regPath -Name "LoadBehavior" -Value 3 -Type DWord -ErrorAction SilentlyContinue

    [WizardClassLibrary.WordApp]::CheckAndRegisterWordAddin($wordAddinDll)
    if (-not [WizardClassLibrary.WordApp]::StartApp()) { throw "FLIR WordApp could not start." }
    
    $word = [WizardClassLibrary.WordApp]::GetWordApp()
    $word.Visible = $true
    $word.DisplayAlerts = 0
    
    try {
        $word.Options.AutoFormatAsYouTypeReplaceText = $false
        $word.AutoCorrect.ReplaceText = $false
        # Disable online web updates/verification for Office COM instance
        $word.Options.WarnBeforeSavingPrintingSendingWithMarkup = $false
    } catch {}

    foreach ($manifest in $manifestArray) {
        # Generate unique temporary files per manifest entry
        $analysisPath  = [System.IO.Path]::GetTempFileName()
        $reportXmlPath = [System.IO.Path]::GetTempFileName()
        $document      = $null

        try {
            $templatePath = [string]$manifest.template_path
            if (-not (Test-Path $templatePath)) {
                throw "Template document not found: $templatePath"
            }

            # Validate IR path if present in manifest
            $firstPage = $manifest.pages[0]
            if ($firstPage -and -not [string]::IsNullOrWhiteSpace($firstPage.ir_path)) {
                if (-not (Test-Path $firstPage.ir_path)) {
                    Write-Warning "IR Image path not found: $($firstPage.ir_path). Skipping FLIR IR injection for this page."
                }
            }

            # FLIR Template Analysis & XML Generation
            [WizardClassLibrary.WordApp]::AnalyzeTemplate($templatePath, $analysisPath, [ref]$false) | Out-Null
            
            $report = New-ProReport -Manifest $manifest -Word $word
            Serialize-ProReport -Report $report -OutputPath $reportXmlPath        
            
            # Execute FLIR SDK Report Creation (keep Word hidden)
            $word.Visible = $false
            [WizardClassLibrary.WordApp]::CreateReport($reportXmlPath, [ref]$false) | Out-Null
            $word.Visible = $false

            # Resolve active document after FLIR SDK engine pass
            if ($null -ne $word.ActiveDocument) {
                $document = $word.ActiveDocument
            } elseif ($word.Documents.Count -gt 0) {
                $document = $word.Documents.Item($word.Documents.Count)
            } else {
                throw "Word did not keep any document open after CreateReport."
            }

            # Perform Word manipulations only if document handle is valid
            if ($null -eq $document) {
                throw "Active document reference is null."
            }

            Apply-FieldOverrides -Document $document -Manifest $manifest
            Replace-TextWithImage -Document $document -Manifest $manifest
            Remove-EmptyParagraphsInHeadersAndFooters -Document $document
            Remove-PageBreaksAndMerge -Document $document
            Remove-EmptyParagraphsAtEdges -Document $document

            # Save generated output
            $outputPath = [string]$manifest.output_path
            $outputDir  = Split-Path -Path $outputPath -Parent
            if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir -Force | Out-Null }
            if (Test-Path $outputPath) { Remove-Item -Path $outputPath -Force }

            $document.SaveAs2($outputPath)
            Write-Output "Successfully Generated: $outputPath"

            $document.Close($false)
            $document = $null

        } catch {
            Write-Error "Error processing $($manifest.output_path): $($_.Exception.Message)"
        } finally {
            if ($null -ne $document) {
                try { $document.Close($false) } catch {}
                $document = $null
            }
            if ($analysisPath -and (Test-Path $analysisPath)) {
                Remove-Item $analysisPath -Force -ErrorAction SilentlyContinue
            }
            if ($reportXmlPath -and (Test-Path $reportXmlPath)) {
                Remove-Item $reportXmlPath -Force -ErrorAction SilentlyContinue
            }
        }
    }
} catch {
    Write-Error $_.Exception.Message
    exit 1
} finally {
    try { [WizardClassLibrary.WordApp]::CloseWordApp() } catch {}
    
    if ($null -ne $word) {
        try { 
            $word.Quit($false) 
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
        } catch {}
    }
    
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}