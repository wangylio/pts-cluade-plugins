param(
    [string]$Root = '.',
    [int]$Index = 0,
    [string]$OutDir = (Join-Path $env:TEMP 'tsd_read'),
    [string]$Extensions = '.doc,.docx,.xls,.xlsx,.ppt,.pptx',
    [switch]$List
)

$ErrorActionPreference = 'Stop'
$exts = @($Extensions.Split(',') | ForEach-Object { $_.Trim().ToLower() })

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
        $text = [string]$doc.Content.Text
    }
    elseif ($ext -eq '.xls' -or $ext -eq '.xlsx') {
        $app = New-Object -ComObject Ket.Application
        try { $app.Visible = $false } catch {}
        try { $app.DisplayAlerts = $false } catch {}
        $doc = $app.Workbooks.Open($file.FullName, 0, $true)
        $sb = New-Object System.Text.StringBuilder
        foreach ($ws in $doc.Worksheets) {
            [void]$sb.AppendLine('=== SHEET: ' + $ws.Name + ' ===')
            $ur = $ws.UsedRange
            if ($ur -ne $null) {
                $rows = [int]$ur.Rows.Count
                $cols = [int]$ur.Columns.Count
                $vals = $null
                try { $vals = $ur.Value2 } catch { }
                for ($r = 1; $r -le $rows; $r++) {
                    $line = @()
                    for ($c = 1; $c -le $cols; $c++) {
                        $cell = $vals
                        if ($vals -is [System.Array]) { $cell = $vals[($r - 1), ($c - 1)] }
                        if ($cell -is [System.DBNull]) { $cell = '' }
                        $line += [string]$cell
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
            foreach ($sh in $s.Shapes) {
                try {
                    if ($sh.HasTextFrame -and $sh.TextFrame.HasText) {
                        [void]$sb.AppendLine([string]$sh.TextFrame.TextRange.Text)
                    }
                } catch {}
            }
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
    $out = Join-Path $OutDir ('tsd_{0}.txt' -f $Index)
    Set-Content -Path $out -Value $text -Encoding UTF8
    Write-Output ('OK chars={0} -> {1}' -f $text.Length, $out)
}
