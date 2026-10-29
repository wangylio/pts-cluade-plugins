param(
    [string]$Root = '.',
    [int]$Index = 0,
    [string]$OutDir = (Join-Path $env:TEMP 'tsd_read'),
    [string]$Extensions = '.doc,.docx,.xls,.xlsx,.ppt,.pptx',
    [switch]$List,
    [switch]$Cleanup
)

$ErrorActionPreference = 'Stop'
$exts = @($Extensions.Split(',') | ForEach-Object { $_.Trim().ToLower() })

if ($Cleanup) {
    if (Test-Path $OutDir) {
        Remove-Item -Recurse -Force $OutDir
        Write-Output ('CLEANED: ' + $OutDir)
    } else {
        Write-Output ('CLEANED: (absent) ' + $OutDir)
    }
    exit 0
}

$CR = [string][char]13
$BEL = [string][char]7     # end-of-cell / end-of-row mark
$VT = [string][char]11      # manual line break
$FF = [string][char]12      # page / section break
$SOH = [string][char]1      # inline object anchor
$STX = [string][char]2      # footnote / endnote reference
$FC1 = [string][char]19     # field begin
$FC2 = [string][char]20     # field separator
$FC3 = [string][char]21     # field end

function Test-Tsd([string]$p) {
    try {
        $fs = [System.IO.File]::OpenRead($p)
        $buf = New-Object byte[] 16
        $n = $fs.Read($buf, 0, 16)
        $fs.Close()
        if ($n -lt 12) { return $false }
        $s = [System.Text.Encoding]::ASCII.GetString($buf)
        return $s.StartsWith('%TSD-Header')
    } catch { return $false }
}

# Word text keeps table cells as CR+BEL and rows as CR+BEL+CR+BEL.
# Convert cells to tabs and rows to newlines so structure survives.
function Convert-WordText([string]$raw) {
    $t = $raw.Replace($CR + $BEL + $CR + $BEL, "`n")
    $t = $t.Replace($CR + $BEL, "`t")
    $t = $t.Replace($BEL, "`n")
    $t = $t.Replace($VT, "`n")
    $t = $t.Replace($FF, "`n[PAGE BREAK]`n")
    $t = $t.Replace($SOH, '[OBJECT]')
    $t = $t.Replace($STX, '')
    $t = $t.Replace($FC1, '').Replace($FC2, '').Replace($FC3, '')
    $t = $t.Replace($CR, "`n")
    return $t
}

function Add-StoryText($list, $collection, [string]$label) {
    foreach ($i in 1, 2, 3) {
        $txt = ''
        try { $txt = ([string]$collection.Item($i).Range.Text).Trim() } catch {}
        if ($txt -ne '') {
            $line = '[' + $label + '] ' + $txt
            if (-not $list.Contains($line)) { [void]$list.Add($line) }
        }
    }
}

function Add-ItemTexts($list, $collection, [string]$label) {
    $n = 0
    try { $n = [int]$collection.Count } catch { return }
    if ($n -gt 500) { $n = 500 }
    for ($i = 1; $i -le $n; $i++) {
        $txt = ''
        try { $txt = ([string]$collection.Item($i).Range.Text).Trim() } catch {}
        if ($txt -ne '') {
            $line = '[' + $label + ' ' + $i + '] ' + $txt
            if (-not $list.Contains($line)) { [void]$list.Add($line) }
        }
    }
}

# Heading outline, driven by OutlineLevel so it is language independent.
function Get-Outline($doc) {
    $list = New-Object System.Collections.Generic.List[string]
    try { $null = $doc.Paragraphs.Item(1).OutlineLevel } catch { return $list }
    $n = 0
    try { $n = [int]$doc.Paragraphs.Count } catch { return $list }
    if ($n -gt 20000) { $n = 20000 }
    for ($i = 1; $i -le $n; $i++) {
        $lv = 10
        $txt = ''
        try {
            $p = $doc.Paragraphs.Item($i)
            $lv = [int]$p.OutlineLevel
            if ($lv -ge 1 -and $lv -le 9) { $txt = ([string]$p.Range.Text).Trim() }
        } catch { continue }
        if ($txt -ne '') {
            $pad = ''
            for ($k = 1; $k -lt $lv; $k++) { $pad += '  ' }
            [void]$list.Add($pad + 'H' + $lv + ': ' + $txt)
        }
    }
    return $list
}

# Bulk value / format reads may come back as a 1-D array for single-row or
# single-column ranges, so index according to the actual array rank.
# WPS returns COM SAFEARRAYs with a lower bound of 1, so index through
# GetValue() with the real lower bound instead of PowerShell's [..] syntax.
function Get-ItemByPos($arr, $r, $c, $rows, $cols) {
    if ($null -eq $arr) { return $null }
    if ($arr -is [System.Array]) {
        try {
            $lb0 = $arr.GetLowerBound(0)
            if ($arr.Rank -eq 2) {
                return $arr.GetValue($lb0 + $r - 1, $arr.GetLowerBound(1) + $c - 1)
            }
            if ($rows -eq 1) { return $arr.GetValue($lb0 + $c - 1) }
            if ($cols -eq 1) { return $arr.GetValue($lb0 + $r - 1) }
            return $arr.GetValue($lb0 + ($r - 1) * $cols + ($c - 1))
        } catch { return $null }
    }
    if ($arr -is [System.DBNull]) { return $null }
    return $arr
}

function Test-DateFormat([string]$f) {
    if ([string]::IsNullOrEmpty($f)) { return $false }
    $s = $f -replace '"[^"]*"', ''
    $s = $s -replace '\[[^\]]*\]', ''
    $hasY = $s -match '[yY]'
    $hasD = $s -match '[dD]'
    $hasH = $s -match '[hH]'
    $hasS = $s -match '[sS]'
    if ($hasY -and $hasD) { return $true }
    if (($hasH -or $hasS) -and $s.Contains(':')) { return $true }
    return $false
}

# Quote cells that contain tabs, newlines or quotes (TSV convention).
function Format-Cell($v) {
    if ($null -eq $v) { return '' }
    if ($v -is [System.DBNull]) { return '' }
    $s = [string]$v
    $s = $s.Replace($CR + "`n", "`n").Replace($CR, "`n")
    if ($s.Contains("`t") -or $s.Contains("`n") -or $s.Contains('"')) {
        $s = '"' + $s.Replace('"', '""') + '"'
    }
    return $s
}

function Get-ShapeTexts($shapes, $sb, [int]$depth) {
    if ($depth -gt 8) { return }
    foreach ($sh in $shapes) {
        $ty = 0
        try { $ty = [int]$sh.Type } catch {}
        if ($ty -eq 6) {
            try { Get-ShapeTexts $sh.GroupItems $sb ($depth + 1) } catch {}
            continue
        }
        $isTable = $false
        try { $isTable = [bool]$sh.HasTable } catch {}
        if ($isTable) {
            $nr = 0
            $nc = 0
            try {
                $nr = [int]$sh.Table.Rows.Count
                $nc = [int]$sh.Table.Columns.Count
            } catch {}
            if ($nr -gt 0 -and $nc -gt 0) {
                [void]$sb.AppendLine('[TABLE]')
                for ($r = 1; $r -le $nr; $r++) {
                    $line = @()
                    for ($c = 1; $c -le $nc; $c++) {
                        $t = ''
                        try {
                            $tf = $sh.Table.Cell($r, $c).Shape.TextFrame
                            if ($tf.HasText) { $t = [string]$tf.TextRange.Text }
                        } catch {}
                        $t = $t.Replace($CR, ' ').Replace("`n", ' ').Trim()
                        $line += $t
                    }
                    [void]$sb.AppendLine(($line -join "`t"))
                }
                [void]$sb.AppendLine('[/TABLE]')
            }
            continue
        }
        try {
            if ($sh.HasTextFrame -and $sh.TextFrame.HasText) {
                $t = [string]$sh.TextFrame.TextRange.Text
                $t = $t.Replace($CR, "`n").Trim()
                if ($t -ne '') { [void]$sb.AppendLine($t) }
            }
        } catch {}
    }
}

$files = @(Get-ChildItem -Path $Root -Recurse -File -ErrorAction SilentlyContinue |
    Where-Object { $exts -contains $_.Extension.ToLower() } |
    Where-Object { Test-Tsd $_.FullName })

if ($List -or $Index -le 0) {
    $i = 0
    foreach ($f in $files) {
        $i++
        Write-Output ('[{0}] {1} ({2} bytes)' -f $i, $f.FullName, $f.Length)
    }
    Write-Output ('TOTAL: ' + $files.Count)
    exit 0
}

if ($Index -gt $files.Count) { Write-Output 'INDEX-OUT-OF-RANGE'; exit 1 }
$file = $files[$Index - 1]
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$ext = $file.Extension.ToLower()

$app = $null
$doc = $null
$text = ''
$ok = $false

try {
    if ($ext -eq '.doc' -or $ext -eq '.docx') {
        $app = New-Object -ComObject KWps.Application
        try { $app.Visible = $false } catch {}
        try { $app.DisplayAlerts = 0 } catch {}
        $doc = $app.Documents.Open($file.FullName, $false, $true)
        $sb = New-Object System.Text.StringBuilder
        $body = Convert-WordText ([string]$doc.Content.Text)
        [void]$sb.AppendLine($body)

        $outline = Get-Outline $doc
        if ($outline.Count -gt 0) {
            [void]$sb.AppendLine('=== OUTLINE ===')
            foreach ($l in $outline) { [void]$sb.AppendLine($l) }
        }

        $extras = New-Object System.Collections.Generic.List[string]
        try {
            foreach ($sec in $doc.Sections) {
                Add-StoryText $extras $sec.Headers 'header'
                Add-StoryText $extras $sec.Footers 'footer'
            }
        } catch {}
        Add-ItemTexts $extras $doc.Footnotes 'footnote'
        Add-ItemTexts $extras $doc.Endnotes 'endnote'
        Add-ItemTexts $extras $doc.Comments 'comment'
        if ($extras.Count -gt 0) {
            [void]$sb.AppendLine('=== HEADER / FOOTER / FOOTNOTE / COMMENT ===')
            foreach ($l in $extras) { [void]$sb.AppendLine($l) }
        }
        $text = $sb.ToString()
    }
    elseif ($ext -eq '.xls' -or $ext -eq '.xlsx') {
        $app = New-Object -ComObject Ket.Application
        try { $app.Visible = $false } catch {}
        try { $app.DisplayAlerts = $false } catch {}
        $doc = $app.Workbooks.Open($file.FullName, 0, $true)
        $sb = New-Object System.Text.StringBuilder
        foreach ($ws in $doc.Worksheets) {
            $vis = ''
            try { if ([int]$ws.Visible -ne -1) { $vis = ' (hidden)' } } catch {}
            [void]$sb.AppendLine('=== SHEET: ' + $ws.Name + $vis + ' ===')
            $ur = $ws.UsedRange
            if ($ur -ne $null) {
                $rows = [int]$ur.Rows.Count
                $cols = [int]$ur.Columns.Count
                $vals = $null
                try { $vals = $ur.Value2 } catch {}
                $fmts = $null
                $bulkFmt = $false
                try {
                    $fmts = $ur.NumberFormat
                    if ($null -ne $fmts -and -not ($fmts -is [System.DBNull])) { $bulkFmt = $true }
                } catch {}
                for ($r = 1; $r -le $rows; $r++) {
                    $line = @()
                    for ($c = 1; $c -le $cols; $c++) {
                        $cell = Get-ItemByPos $vals $r $c $rows $cols
                        $isNum = $false
                        if ($null -ne $cell -and -not ($cell -is [System.DBNull])) {
                            $isNum = ($cell -is [double]) -or ($cell -is [int]) -or
                                     ($cell -is [long]) -or ($cell -is [decimal])
                        }
                        if ($isNum) {
                            # A range with mixed formats returns null in bulk, so
                            # fall back to reading the format cell by cell.
                            $fs = ''
                            if ($bulkFmt) { $fs = [string](Get-ItemByPos $fmts $r $c $rows $cols) }
                            if ($fs -eq '') {
                                try { $fs = [string]$ur.Cells.Item($r, $c).NumberFormat } catch {}
                            }
                            if (Test-DateFormat $fs) {
                                $d = [double]$cell
                                if ($d -gt 0 -and $d -lt 2958466) {
                                    try {
                                        $dt = [DateTime]::FromOADate($d)
                                        $sl = $fs.ToLower()
                                        if ($sl -match '[yYdD]') {
                                            $cell = $dt.ToString('yyyy-MM-dd HH:mm:ss')
                                        } else {
                                            $cell = $dt.ToString('HH:mm:ss')
                                        }
                                    } catch {}
                                }
                            }
                        }
                        $line += (Format-Cell $cell)
                    }
                    if (($line -join '').Trim() -ne '') { [void]$sb.AppendLine(($line -join "`t")) }
                }
            }
        }
        $text = $sb.ToString()
    }
    elseif ($ext -eq '.ppt' -or $ext -eq '.pptx') {
        $app = New-Object -ComObject KWPP.Application
        try { $app.Visible = $false } catch {}
        $doc = $app.Presentations.Open($file.FullName, $true, $false, $false)
        $sb = New-Object System.Text.StringBuilder
        $n = 1
        foreach ($s in $doc.Slides) {
            [void]$sb.AppendLine('=== SLIDE ' + $n + ' ===')
            try { Get-ShapeTexts $s.Shapes $sb 0 } catch {}
            try {
                foreach ($sh in $s.NotesPage.Shapes) {
                    try {
                        if ($sh.HasTextFrame -and $sh.TextFrame.HasText) {
                            $t = ([string]$sh.TextFrame.TextRange.Text).Replace($CR, "`n").Trim()
                            if ($t -ne '') { [void]$sb.AppendLine('[NOTES] ' + $t) }
                        }
                    } catch {}
                }
            } catch {}
            $n++
        }
        $text = $sb.ToString()
    }
    else {
        Write-Output 'UNSUPPORTED-EXT'
        exit 1
    }
    $ok = $true
} catch {
    Write-Output ('ERR: ' + $_.Exception.Message)
} finally {
    if ($ok -and $doc) { try { $doc.Close(0) } catch {} }
    if ($app) {
        try { $app.Quit() } catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } catch {}
    }
}

if ($ok) {
    $header = '=== SOURCE: ' + $file.FullName + ' (' + $ext + ') ==='
    $out = Join-Path $OutDir ('tsd_{0}.txt' -f $Index)
    Set-Content -Path $out -Value ($header + "`r`n" + $text) -Encoding UTF8
    Write-Output ('OK chars={0} -> {1}' -f $text.Length, $out)
}
