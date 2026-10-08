"""Lint a canvas scene, fix it, export before/after .pptx, round-trip check, and (on Windows with PowerPoint) render + measure.

    python demo.py                      # lint + fix + export + round trip
    python demo.py --render             # also render PNGs and measure text in real PowerPoint (Windows only)
"""
from __future__ import annotations

import argparse
import json
import subprocess
import sys
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

from exportlint.core import export, fits, fix, lint, load, round_trip

HERE = Path(__file__).parent
OUT = HERE / "out"


def ps(script: str, *args: str) -> None:
    subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(HERE / "exportlint" / script), *args], check=True)


def agreement(scene: dict, measured: Path) -> list[tuple[str, bool, bool, int, int]]:
    """Per text layer: (name, predicted fits, PowerPoint fits, predicted lines, PowerPoint lines)."""
    m = {r["name"]: r for r in json.loads(measured.read_text(encoding="utf-8-sig"))}
    rows = []
    for s in scene["slides"]:
        for L in s["layers"]:
            if L["type"] == "text":
                ok, n, _ = fits(L)
                r = m[f"{s['id']}/{L['id']}"]
                rows.append((f"{s['id']}/{L['id']}", ok, r["text_height_pt"] + r["margins_pt"] <= r["box_height_pt"] + 0.5, n, r["lines"]))
    return rows


def side_by_side(n_slides: int) -> Path:
    label = ImageFont.truetype("arialbd.ttf", 34)
    rows = []
    for i in range(1, n_slides + 1):
        a, b = (Image.open(OUT / "png" / f"{t}-{i}.png").convert("RGB") for t in ("before", "after"))
        row = Image.new("RGB", (a.width + b.width + 40, a.height + 60), "white")
        row.paste(a, (0, 60)), row.paste(b, (a.width + 40, 60))
        d = ImageDraw.Draw(row)
        d.text((10, 10), f"Slide {i} — as exported today", fill="#B00020", font=label)
        d.text((a.width + 50, 10), f"Slide {i} — after exportlint fixes", fill="#0A7A3E", font=label)
        rows.append(row)
    sheet = Image.new("RGB", (rows[0].width, sum(r.height for r in rows)), "white")
    y = 0
    for r in rows:
        sheet.paste(r, (0, y))
        y += r.height
    sheet = sheet.resize((sheet.width // 2, sheet.height // 2))
    path = HERE / "docs" / "before_after.png"
    path.parent.mkdir(exist_ok=True)
    sheet.save(path)
    return path


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("scene", nargs="?", default=str(HERE / "samples" / "northstar_deck.json"))
    ap.add_argument("--render", action="store_true", help="render + measure in local PowerPoint via COM (Windows)")
    a = ap.parse_args()
    OUT.mkdir(exist_ok=True)

    scene = load(Path(a.scene))
    issues = lint(scene)
    print(f"\n== {scene['name']}: {len(issues)} export-fidelity issues before export\n")
    for i in issues:
        print(f"  [{i.code:<18}] {i.slide}/{i.layer}: {i.detail}")
        if i.fix:
            print(f"  {'':20} fix → {i.fix}")

    fixed, log = fix(scene)
    print(f"\n== auto-fixes applied ({len(log)})\n")
    for line in log:
        print("  " + line)
    left = lint(fixed)
    print(f"\n== issues after fixing: {len(left)}")

    export(scene, OUT / "before.pptx")
    export(fixed, OUT / "after.pptx")
    drift = round_trip(scene, OUT / "before.pptx") + round_trip(fixed, OUT / "after.pptx")
    print(f"== round trip (.pptx re-read): {len(drift)} layers lost / moved / changed")
    for d in drift:
        print(f"  [{d.code}] {d.slide}/{d.layer}: {d.detail}")

    if a.render:
        for tag in ("before", "after"):
            ps("render.ps1", "-Pptx", str(OUT / f"{tag}.pptx"), "-OutDir", str(OUT / "png"), "-Prefix", tag)
            ps("measure.ps1", "-Pptx", str(OUT / f"{tag}.pptx"), "-Out", str(OUT / f"measured_{tag}.json"))
        print(f"== rendered in PowerPoint → {side_by_side(len(scene['slides']))}")

    for tag, sc in (("before", scene), ("after", fixed)):
        mp = OUT / f"measured_{tag}.json"
        if mp.exists():
            rows = agreement(sc, mp)
            fit_ok = sum(p == q for _, p, q, _, _ in rows)
            lines_ok = sum(x == y for *_, x, y in rows)
            over = [n for n, _, q, _, _ in rows if not q]
            print(f"== PowerPoint check ({tag}): overflow prediction {fit_ok}/{len(rows)}, line count {lines_ok}/{len(rows)}; "
                  f"overflowing in PowerPoint: {', '.join(over) or 'none'}")
    print("\nfiles: out/before.pptx  out/after.pptx" + ("  docs/before_after.png" if a.render else ""))
    return 0 if not left and not drift else 1


if __name__ == "__main__":
    sys.exit(main())
