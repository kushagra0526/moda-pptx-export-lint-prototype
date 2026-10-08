# Ask PowerPoint itself how tall each text box's text really is (ground truth for the overflow predictor).
# usage: powershell -ExecutionPolicy Bypass -File measure.ps1 -Pptx deck.pptx -Out measured.json
param([string]$Pptx, [string]$Out)
$Pptx = (Resolve-Path $Pptx).Path
$app = New-Object -ComObject PowerPoint.Application
$rows = @()
try {
    $p = $app.Presentations.Open($Pptx, -1, 0, 0)
    foreach ($s in $p.Slides) {
        foreach ($sh in $s.Shapes) {
            if ($sh.HasTextFrame -and $sh.TextFrame.HasText) {
                $tr = $sh.TextFrame.TextRange
                $rows += [pscustomobject]@{ name = $sh.Name; lines = $tr.Lines().Count; text_height_pt = $tr.BoundHeight;
                                            box_height_pt = $sh.Height; margins_pt = $sh.TextFrame.MarginTop + $sh.TextFrame.MarginBottom }
            }
        }
    }
    $p.Close()
} finally {
    $app.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
}
$rows | ConvertTo-Json | Set-Content -Encoding utf8 $Out
