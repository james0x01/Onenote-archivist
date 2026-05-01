# export_onenote_hierarchy.ps1
# Extracts exact page hierarchy from OneNote desktop via COM API.
# Writes rollup_groups.json per notebook -- 100% accurate, no LLM needed.
#
# Requirements: Windows + OneNote desktop + notebooks synced
# No extra installs needed -- uses built-in PowerShell COM support.
#
# Usage:
#   .\export_onenote_hierarchy.ps1
#   .\export_onenote_hierarchy.ps1 -OutputDir C:\onenote_export
#   .\export_onenote_hierarchy.ps1 -FilterNotebook "Training"
#
# After running, copy each rollup_groups.json to Mac Dropbox at:
#   01_Raw_Audit/{Notebook}/rollup_groups.json
# Then SCP to Cerebro and run: python3 summarize_rollups.py --force

param(
    [string]$OutputDir      = ".\rollup_groups",
    [string]$FilterNotebook = ""
)

Set-StrictMode -Off
$ErrorActionPreference = "Stop"

$SkipSections = @("candidates")
$MinSectionSz = 4

# ---------------------------------------------------------------------------
# Helper: strip characters that are illegal in folder/file names
# ---------------------------------------------------------------------------
function Remove-InvalidChars {
    param([string]$Name)
    $result = $Name
    foreach ($c in [System.IO.Path]::GetInvalidFileNameChars()) {
        $result = $result.Replace([string]$c, '')
    }
    return $result.Trim()
}

# ---------------------------------------------------------------------------
# Banner
# ---------------------------------------------------------------------------
Write-Host ("=" * 60)
Write-Host "  export_onenote_hierarchy.ps1 -- OneNote COM export"
Write-Host ("=" * 60)
Write-Host ""

# ---------------------------------------------------------------------------
# Connect to OneNote
# ---------------------------------------------------------------------------
Write-Host "Connecting to OneNote desktop..."
try {
    $onenote = New-Object -ComObject "OneNote.Application"
} catch {
    Write-Error "Could not connect to OneNote. Make sure OneNote desktop is installed and has been opened at least once."
    exit 1
}

Write-Host "Fetching full hierarchy (may take a moment)..."
$xmlStr = ""
$onenote.GetHierarchy("", 4, [ref]$xmlStr)

$doc = [xml]$xmlStr
New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null

$totalNb     = 0
$totalGroups = 0

# ---------------------------------------------------------------------------
# Walk notebooks
# ---------------------------------------------------------------------------
foreach ($nb in $doc.DocumentElement.ChildNodes) {
    if ($nb.LocalName -ne "Notebook") { continue }

    $nbName = $nb.GetAttribute("name")
    if ($FilterNotebook -and $nbName -ne $FilterNotebook) { continue }

    Write-Host ""
    Write-Host "Notebook: $nbName"

    $nbFolder = Remove-InvalidChars $nbName
    $config   = [ordered]@{}

    $allNodes = $nb.SelectNodes(".//*")
    foreach ($sec in $allNodes) {
        if ($sec.LocalName -ne "Section") { continue }

        $secName = $sec.GetAttribute("name")
        if ([string]::IsNullOrEmpty($secName))          { continue }
        if ($SkipSections -contains $secName.ToLower()) { continue }

        # Collect page nodes
        $pageNodes = @()
        foreach ($child in $sec.ChildNodes) {
            if ($child.LocalName -eq "Page") {
                $pageNodes += $child
            }
        }
        if ($pageNodes.Count -lt $MinSectionSz) { continue }

        # Build page list with 0-based level
        $pages = @()
        foreach ($pg in $pageNodes) {
            $lvlStr = $pg.GetAttribute("pageLevel")
            $lvl    = 0
            if (-not [string]::IsNullOrEmpty($lvlStr)) {
                $lvl = [int]$lvlStr - 1
            }
            $obj = New-Object PSObject -Property @{
                Name  = $pg.GetAttribute("name")
                Level = $lvl
            }
            $pages += $obj
        }

        # Group level-1+ pages under the preceding level-0 parent
        $groups  = [ordered]@{}
        $other   = New-Object System.Collections.Generic.List[string]
        $current = $null

        foreach ($p in $pages) {
            if ($p.Level -eq 0) {
                $current = $p.Name
                if (-not $groups.Contains($current)) {
                    $groups[$current] = New-Object System.Collections.Generic.List[string]
                }
            } else {
                if ($null -ne $current) {
                    $groups[$current].Add($p.Name)
                } else {
                    $other.Add($p.Name)
                }
            }
        }

        # Move childless level-0 pages to Other
        $childlessKeys = @()
        foreach ($k in $groups.Keys) {
            if ($groups[$k].Count -eq 0) {
                $childlessKeys += $k
            }
        }
        foreach ($k in $childlessKeys) {
            $other.Add($k)
            $groups.Remove($k)
        }

        if ($other.Count -gt 0) { $groups["Other"] = $other }
        if ($groups.Count -eq 0) { continue }

        # Convert lists to plain arrays for clean JSON output
        $clean = [ordered]@{}
        foreach ($gName in $groups.Keys) {
            $clean[$gName] = @($groups[$gName])
        }

        $config[$secName] = $clean

        $nOther  = 0
        if ($clean.Contains("Other")) { $nOther = $clean["Other"].Count }
        $nGroups = 0
        foreach ($k in $clean.Keys) {
            if ($k -ne "Other") { $nGroups++ }
        }

        Write-Host ("  {0}: {1} group(s), {2} ungrouped" -f $secName, $nGroups, $nOther)
        $totalGroups += $nGroups
    }

    if ($config.Count -eq 0) {
        Write-Host "  (no sections with hierarchy data)"
        continue
    }

    $nbDir   = Join-Path $OutputDir $nbFolder
    New-Item -ItemType Directory -Force -Path $nbDir | Out-Null
    $outFile = Join-Path $nbDir "rollup_groups.json"
    $json    = $config | ConvertTo-Json -Depth 10
    [System.IO.File]::WriteAllText($outFile, $json, (New-Object System.Text.UTF8Encoding($false)))
    Write-Host "  Written: $outFile"
    $totalNb++
}

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------
Write-Host ""
Write-Host ("=" * 60)
Write-Host "  EXPORT COMPLETE"
Write-Host ("  Notebooks : {0}" -f $totalNb)
Write-Host ("  Groups    : {0}" -f $totalGroups)
Write-Host ("  Output    : {0}" -f (Resolve-Path $OutputDir))
Write-Host ("=" * 60)
Write-Host ""
Write-Host "Next steps:"
Write-Host "  1. Copy each rollup_groups.json to Mac Dropbox:"
Write-Host "       01_Raw_Audit/{Notebook}/rollup_groups.json"
Write-Host ""
Write-Host "  2. SCP each file to Cerebro (run from Mac terminal):"
Write-Host "       scp <Dropbox>/01_Raw_Audit/<Notebook>/rollup_groups.json lab-user@10.254.254.48:Onenote-archivist/onenote_audit/01_Raw_Audit/<Notebook>/"
Write-Host ""
Write-Host "  3. On Cerebro:"
Write-Host "       python3 summarize_rollups.py --force"
