# Render every slide of a .pptx to PNG with the locally installed PowerPoint (Windows, COM).
# usage: powershell -ExecutionPolicy Bypass -File render.ps1 -Pptx deck.pptx -OutDir out\png -Prefix before
param([string]$Pptx, [string]$OutDir, [string]$Prefix = "slide")
$Pptx = (Resolve-Path $Pptx).Path
New-Item -ItemType Directory -Force $OutDir | Out-Null
$OutDir = (Resolve-Path $OutDir).Path
$app = New-Object -ComObject PowerPoint.Application
try {
    # ReadOnly, Untitled, WithWindow = false
    $p = $app.Presentations.Open($Pptx, -1, 0, 0)
    for ($i = 1; $i -le $p.Slides.Count; $i++) {
        $p.Slides.Item($i).Export((Join-Path $OutDir "$Prefix-$i.png"), "PNG", 1600, 900)
    }
    $p.Close()
} finally {
    $app.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
}
