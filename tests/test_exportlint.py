import copy
from pathlib import Path

import pytest

from demo import agreement
from exportlint.core import delta_e, export, fits, fix, font, lint, load, nearest_brand, round_trip, wrap

HERE = Path(__file__).parent
SCENE = load(HERE.parent / "samples" / "northstar_deck.json")


def codes(issues):
    return sorted((i.code, i.layer) for i in issues)


def test_finds_every_planted_issue():
    got = codes(lint(SCENE))
    for want in [("UNSUPPORTED_EFFECT", "orb"), ("UNSUPPORTED_EFFECT", "card"), ("TEXT_OVERFLOW", "title"),
                 ("TEXT_OVERFLOW", "body"), ("TEXT_TOO_SMALL", "foot"), ("OFF_BRAND_COLOR", "m1")]:
        assert want in got
    assert sum(c == "FONT_SUBSTITUTED" for c, _ in got) == 2  # grouped per family, not per layer


def test_no_false_positives_on_clean_layers():
    flagged = {(i.slide, i.layer) for i in lint(SCENE) if i.code not in ("FONT_SUBSTITUTED",)}
    for clean in [("s1", "subtitle"), ("s2", "heading"), ("s2", "quote"), ("s3", "l1")]:
        assert clean not in flagged


def test_fix_clears_everything_and_keeps_input_untouched():
    before = copy.deepcopy(SCENE)
    fixed, log = fix(SCENE)
    assert lint(fixed) == []
    assert SCENE == before
    assert any("shrink 50 → 45" in l for l in log)


def test_shrink_never_goes_below_brand_minimum():
    s = copy.deepcopy(SCENE)
    s["slides"][1]["layers"][3]["box"][3] = 0.5  # body box far too small
    fixed, log = fix(s)
    assert fixed["slides"][1]["layers"][3]["size"] >= s["brand"]["min_text_pt"]
    assert any("needs a human edit" in l for l in log)


def test_delta_e_and_nearest_brand():
    assert delta_e("#14213D", "#14213D") == 0
    assert delta_e("#000000", "#FFFFFF") == pytest.approx(100, abs=0.5)
    assert nearest_brand("#FF5A1F", SCENE["brand"]["colors"])[0] == "orange"


def test_wrap_respects_width():
    f = font("Arial", 20, False)
    for line in wrap("the quick brown fox jumps over the lazy dog " * 4, f, 300):
        assert f.getlength(line) <= 300 or " " not in line


def test_round_trip_has_no_drift(tmp_path):
    fixed, _ = fix(SCENE)
    export(fixed, tmp_path / "a.pptx")
    assert round_trip(fixed, tmp_path / "a.pptx") == []


def test_round_trip_detects_lost_layer(tmp_path):
    export(SCENE, tmp_path / "a.pptx")
    s = copy.deepcopy(SCENE)
    s["slides"][0]["layers"].append({"id": "ghost", "type": "shape", "shape": "rect", "box": [0, 0, 1, 1], "fill": "#14213D"})
    assert ("LAYER_LOST", "ghost") in codes(round_trip(s, tmp_path / "a.pptx"))


@pytest.mark.parametrize("tag", ["before", "after"])
def test_predictions_match_real_powerpoint(tag):
    """Fixtures were measured by PowerPoint (exportlint/measure.ps1) on the exported decks."""
    sc = SCENE if tag == "before" else fix(SCENE)[0]
    rows = agreement(sc, HERE / "fixtures" / f"measured_{tag}.json")
    assert all(p == q for _, p, q, _, _ in rows)
    assert all(x == y for *_, x, y in rows)
