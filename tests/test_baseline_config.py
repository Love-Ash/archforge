# -*- coding: utf-8 -*-
"""Baseline, config, severity, and threshold contracts (split from test_lint.py, 0.8.x)."""
from helpers import *   # noqa: F401,F403
from helpers import (_by, _add_fld, _repo_root, _require_repo_file,
                     patch_theme_fonts)   # noqa: F401


def test_cli_skip_codes(tmp_path):
    """--skip: selectively suppress genre-independent warnings (W14 etc.)."""
    p = new_prs()
    for t in ("시장 현황", "경쟁 구도 분석", "제품 라인업 개요", "사업 확장 전략",
              "재무 운용 계획", "향후 추진 방안"):
        s = add_slide(p)
        tb(s, 1, 0.8, 9, 0.8, t, font="Wanted Sans", size=26)
        tb(s, 1, 2.2, 10, 3, "본문 내용", font="Wanted Sans", size=12)
    path = save(p, tmp_path, "fx.pptx")
    doc = json.loads(run_cli([path, "--json", "--profile", "full"]).stdout)
    assert any(w["code"] == "W14" for w in doc["warnings"])
    doc = json.loads(run_cli([path, "--json", "--profile", "full", "--skip", "W14"]).stdout)
    assert not any(w["code"] == "W14" for w in doc["warnings"])


# ---------------------------------------------------------------- 0.2.1 script layer + fixes

def test_skip_rejects_error_codes(tmp_path):
    """--skip is WARN-only: an E code causes exit 2 rejection, and applied skips
    are recorded in the JSON (fixes a footgun)."""
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "모노 폴백 한글", font="IBM Plex Mono", size=12)
    path = save(p, tmp_path, "fx.pptx")
    r = run_cli([path, "--skip", "E1"])
    assert r.returncode == 2 and "WARN" in r.stderr
    doc = json.loads(run_cli([path, "--json", "--profile", "full", "--skip", "W14"]).stdout)
    assert doc["summary"]["skipped_codes"] == ["W14"]
    assert any(e["code"] == "E1" for e in doc["errors"])   # E1 is still alive

def test_profile_core_drops_style_rules(tmp_path):
    """The core profile = objective defects only: even E2 (a stylistic ERROR) is
    excluded, but the choice is left in the JSON."""
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "이건" + EM_DASH + "차단", font="Wanted Sans", size=12)   # E2 only
    path = save(p, tmp_path, "fx.pptx")
    doc = json.loads(run_cli([path, "--json", "--profile", "full"]).stdout)
    assert any(e["code"] == "E2" for e in doc["errors"])           # blocked when full is specified explicitly (0.4.0: default is core)
    doc = json.loads(run_cli([path, "--json", "--profile", "core"]).stdout)
    assert not doc["errors"] and doc["summary"]["pass"]
    assert doc["summary"]["profile"] == "core"
    assert "E2" in doc["summary"]["skipped_codes"]                  # not a silent bypass

    # objective defects (E1) still get blocked even under core
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "모노 폴백 한글", font="IBM Plex Mono", size=12)
    doc = json.loads(run_cli([save(p, tmp_path, "fx2.pptx"), "--json", "--profile", "core"]).stdout)
    assert any(e["code"] == "E1" for e in doc["errors"])

def test_profile_is_engine_policy(tmp_path, monkeypatch):
    """P0-5: a profile is an execution policy. It can be used as a library via
    lint(profile=), and since an excluded rule simply doesn't run, its internal
    failure doesn't leak into W18."""
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "이건" + EM_DASH + "차단", font="Wanted Sans", size=12)
    path = save(p, tmp_path, "fx.pptx")
    errors, _w = lint_full(path, profile="core")
    assert not by_code(errors, "E2"), errors            # profile works through the library API
    errors, _w = lint_full(path)
    assert by_code(errors, "E2"), errors                # default full stays blocking

    def boom(*a, **kw):
        raise RuntimeError("w9 fail")
    monkeypatch.setattr(jl, "accent_vbars_check", boom)
    _e, warns = lint_full(path, profile="core")           # W9 excluded -> doesn't run -> no W18 leak
    assert not any("w9" in d for (_si, m, d) in by_code(warns, "W18")), warns
    _e, warns = lint_full(path)                           # full -> runs -> guard -> W18
    assert any("w9" in d for (_si, m, d) in by_code(warns, "W18")), warns

def test_skip_validation_strengthened(tmp_path):
    """P1: a nonexistent W code or suppressing W18 causes exit 2 rejection."""
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "clean", font="Wanted Sans", size=20)
    path = save(p, tmp_path, "fx.pptx")
    assert run_cli([path, "--skip", "W51"]).returncode == 2
    assert run_cli([path, "--skip", "W18"]).returncode == 2
    assert run_cli([path, "--skip", "W14"]).returncode == 0

def test_default_profile_is_core(tmp_path):
    """0.4.0 breaking change: with no options, the default is core (objective
    defects only). The style rule E2 is opt-in; the objective defect E1 stays
    blocked even under the default."""
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "이건" + EM_DASH + "정상 통과", font="Wanted Sans", size=12)
    dash_deck = save(p, tmp_path, "dash.pptx")
    errors, _w = jl.lint(dash_deck)              # library default
    assert not by_code(errors, "E2"), errors
    r = run_cli([dash_deck, "--json"])           # CLI default
    doc = json.loads(r.stdout)
    assert r.returncode == 0 and doc["summary"]["profile"] == "core"
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "모노 폴백 한글", font="IBM Plex Mono", size=12)
    assert run_cli([save(p, tmp_path, "e1.pptx")]).returncode == 1   # objective defects stay blocked

def test_config_file(tmp_path):
    """.archforge.json: auto-discovered in the deck folder; a CLI flag beats the
    config."""
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "이건" + EM_DASH + "차단", font="Wanted Sans", size=12)
    deck = save(p, tmp_path, "fx.pptx")
    with open(os.path.join(str(tmp_path), ".archforge.json"), "w", encoding="utf-8") as f:
        json.dump({"profile": "full"}, f)
    doc = json.loads(run_cli([deck, "--json"]).stdout)
    assert any(e["code"] == "E2" for e in doc["errors"])    # the config's full is applied
    doc = json.loads(run_cli([deck, "--json", "--profile", "core"]).stdout)
    assert not doc["errors"]                                 # CLI beats the config
    # fail-safe (fourth review): an unknown key (a typo) causes exit 2 rather than
    # being ignored. Prevents the accident where 'profle: full' silently runs as
    # the default core
    with open(os.path.join(str(tmp_path), ".archforge.json"), "w", encoding="utf-8") as f:
        json.dump({"profile": "full", "no_such_key": 1}, f)
    r = run_cli([deck, "--json"])
    assert r.returncode == 2 and "no_such_key" in r.stderr
    # --no-config: ignore the config of an untrusted deck folder (a trust boundary)
    with open(os.path.join(str(tmp_path), ".archforge.json"), "w", encoding="utf-8") as f:
        json.dump({"profile": "full"}, f)
    doc = json.loads(run_cli([deck, "--json", "--no-config"]).stdout)
    assert not doc["errors"] and doc["summary"]["config"] is None
    doc = json.loads(run_cli([deck, "--json"]).stdout)
    assert doc["summary"]["config"] and doc["summary"]["config"].endswith(".archforge.json")

def test_config_value_validation(tmp_path):
    """A config value's type/range gets a clean exit 2 rather than a traceback
    (fourth review). Same for CLI range (closes the old X1 bypass where
    --hard-min 0 silently turned off E3)."""
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "clean", font="Wanted Sans", size=20)
    deck = save(p, tmp_path, "fx.pptx")
    with open(os.path.join(str(tmp_path), ".archforge.json"), "w", encoding="utf-8") as f:
        json.dump({"hard_min": "abc"}, f)
    r = run_cli([deck, "--json"])
    assert r.returncode == 2 and "hard_min" in r.stderr and "Traceback" not in r.stderr
    with open(os.path.join(str(tmp_path), ".archforge.json"), "w", encoding="utf-8") as f:
        json.dump({"lang": "fr"}, f)
    assert run_cli([deck, "--json"]).returncode == 2
    with open(os.path.join(str(tmp_path), ".archforge.json"), "w", encoding="utf-8") as f:
        json.dump({}, f)
    assert run_cli([deck, "--hard-min", "0"]).returncode == 2   # CLI range validation

def test_baseline_v2_language_and_count(tmp_path):
    """Fingerprint v2 (fourth review HIGH fix): a baseline made under ko is also
    valid under an en run, preserves the meaning of the occurrence count, and is
    page-independent so it survives a slide insertion."""
    p = new_prs()
    for i in range(6):   # a W6 clone deck (the representative rule whose detail used to be a locale string)
        s = add_slide(p)
        tb(s, 1, 0.8, 8, 0.6, "Title block %d" % i, font="Wanted Sans", size=24)
        tb(s, 1, 2.0, 6, 2.5, "Body block", font="Wanted Sans", size=12)
        tb(s, 8, 2.0, 4, 2.5, "Side block", font="Wanted Sans", size=12)
    deck = save(p, tmp_path, "w6.pptx")
    bl = os.path.join(str(tmp_path), "bl.json")
    run_cli([deck, "--profile", "full", "--write-baseline", bl], lang="ko")
    doc = json.loads(run_cli([deck, "--profile", "full", "--json", "--baseline", bl],
                             lang="en").stdout)
    assert not any(w["code"] == "W6" for w in doc["warnings"]), doc["warnings"]

    # count semantics: the same fingerprint occurs 2 times, baseline has 1 -> only 1 is suppressed
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "모노 폴백 한글", font="IBM Plex Mono", size=12)
    one = save(p, tmp_path, "one.pptx")
    run_cli([one, "--write-baseline", bl])
    s = add_slide(p)   # same text, same font = the same fingerprint on two pages
    tb(s, 1, 1, 5, 0.5, "모노 폴백 한글", font="IBM Plex Mono", size=12)
    two = save(p, tmp_path, "two.pptx")
    doc = json.loads(run_cli([two, "--json", "--baseline", bl]).stdout)
    assert doc["summary"]["baseline_suppressed"] == 1
    assert doc["summary"]["error_count"] == 1
    # page-independent: suppression persists even in a deck with a slide inserted before it
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "새로 삽입된 표지", font="Wanted Sans", size=20)
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "모노 폴백 한글", font="IBM Plex Mono", size=12)
    shifted = save(p, tmp_path, "shifted.pptx")
    doc = json.loads(run_cli([shifted, "--json", "--baseline", bl]).stdout)
    assert doc["summary"]["baseline_suppressed"] == 1 and doc["summary"]["error_count"] == 0
    # text-mode visibility: suppression shows as a footnote (fixes the misreading of invisible-as-clean)
    r = run_cli([shifted, "--baseline", bl])
    assert "baseline" in r.stdout

def test_lint_rejects_unknown_profile(tmp_path):
    """Library API: fixes a typo'd profile that used to silently behave as full
    (fourth review)."""
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "clean", font="Wanted Sans", size=20)
    deck = save(p, tmp_path, "fx.pptx")
    with pytest.raises(ValueError):
        jl.lint(deck, profile="ful")

def test_baseline_flow(tmp_path):
    """baseline: accept existing violations, then report only new ones. The
    suppressed count is recorded in summary."""
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "모노 폴백 한글", font="IBM Plex Mono", size=12)
    deck = save(p, tmp_path, "fx.pptx")
    bl = os.path.join(str(tmp_path), "baseline.json")
    r = run_cli([deck, "--write-baseline", bl])
    assert r.returncode == 0 and os.path.exists(bl)
    doc = json.loads(run_cli([deck, "--json", "--baseline", bl]).stdout)
    assert doc["summary"]["pass"] and doc["summary"]["baseline_suppressed"] == 1
    # add a new defect -> only the new one is reported
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "새 결함 한글", font="Consolas", size=12)
    deck2 = save(p, tmp_path, "fx2.pptx")
    doc = json.loads(run_cli([deck2, "--json", "--baseline", bl]).stdout)
    assert doc["summary"]["error_count"] == 1
    assert "Consolas" in doc["errors"][0]["detail"]

def test_nan_thresholds_rejected(tmp_path):
    """NaN slipped through every range comparison (NaN <= 0 is False) and json.load even
    accepts a bare NaN literal, silently disabling E3 via CLI or an attacker-controlled
    config. Both paths must exit 2 now (external verification finding)."""
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 4, 1, "tiny", size=3)
    deck = save(p, tmp_path, "nan.pptx")
    r = run_cli([deck, "--hard-min", "nan"])
    assert r.returncode == 2, (r.returncode, r.stderr)
    cfgdir = os.path.join(str(tmp_path), "cfg")
    os.makedirs(cfgdir)
    shutil.copy(deck, os.path.join(cfgdir, "deck.pptx"))
    with open(os.path.join(cfgdir, ".archforge.json"), "w", encoding="utf-8") as f:
        f.write('{"hard_min": NaN}')
    r = run_cli([os.path.join(cfgdir, "deck.pptx")])
    assert r.returncode == 2, (r.returncode, r.stderr)

def test_scan_rejects_shared_baseline(tmp_path):
    """A single CLI baseline across multiple decks would suppress findings across
    unrelated files (fingerprints carry no file identity): exit 2."""
    d = os.path.join(str(tmp_path), "two")
    os.makedirs(d)
    for name in ("a.pptx", "b.pptx"):
        p = new_prs()
        s = add_slide(p)
        tb(s, 1, 1, 4, 1, "x", size=14)
        save(p, d, name)
    bl = os.path.join(str(tmp_path), "bl.json")
    with open(bl, "w") as f:
        f.write('{"schema_version": "3", "findings": []}')
    r = run_cli(["scan", d, "--baseline", bl])
    assert r.returncode == 2, (r.returncode, r.stdout, r.stderr)

def test_scan_config_lang_does_not_leak(tmp_path):
    """A deck folder config's lang must not change the language of later files or the
    aggregate report (external verification finding: lazy rendering made the last
    file's language dominate everything)."""
    d1 = os.path.join(str(tmp_path), "en_deck")
    d2 = os.path.join(str(tmp_path), "plain")
    os.makedirs(d1)
    os.makedirs(d2)
    for d in (d1, d2):
        p = new_prs()
        s = add_slide(p)
        tb(s, 1, 1, 4, 1, "x", size=14)
        save(p, d, "deck.pptx")
    with open(os.path.join(d1, ".archforge.json"), "w") as f:
        f.write('{"lang": "en"}')
    r = run_cli(["scan", os.path.join(d1, "deck.pptx"), os.path.join(d2, "deck.pptx"),
                 "--json"], lang="ko")
    assert r.returncode == 0, r.stderr
    doc = json.loads(r.stdout)
    assert doc["lang"] == "ko", doc["lang"]

def test_write_baseline_validates_flags_first(tmp_path):
    """A typo'd --skip used to record a baseline as if nothing were wrong: validation
    now precedes recording (external review)."""
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 4, 1, "x", size=14)
    deck = save(p, tmp_path, "v.pptx")
    bl = os.path.join(str(tmp_path), "out_bl.json")
    r = run_cli([deck, "--write-baseline", bl, "--skip", "E1"])
    assert r.returncode == 2
    assert not os.path.exists(bl)
    r = run_cli([deck, "--write-baseline", bl, "--skip", "W999"])
    assert r.returncode == 2
    assert not os.path.exists(bl)

def test_config_no_config_conflict(tmp_path):
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 4, 1, "x", size=14)
    deck = save(p, tmp_path, "c.pptx")
    cfg = os.path.join(str(tmp_path), "x.json")
    with open(cfg, "w") as f:
        f.write("{}")
    r = run_cli([deck, "--config", cfg, "--no-config"])
    assert r.returncode == 2

def test_config_rejects_bool_and_fractional_cluster(tmp_path):
    """float(True) == 1.0 slipped through as a threshold, and w6_cluster 1.9 silently
    truncated to 1 (external review P1)."""
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 4, 1, "x", size=14)
    for body in ('{"hard_min": true}', '{"w6_cluster": 1.9}'):
        d = os.path.join(str(tmp_path), "c%d" % (abs(hash(body)) % 97))
        os.makedirs(d, exist_ok=True)
        deck = save(p, d, "deck.pptx")
        with open(os.path.join(d, ".archforge.json"), "w") as f:
            f.write(body)
        r = run_cli([deck])
        assert r.returncode == 2, (body, r.returncode, r.stderr)

def test_baseline_v3_move_not_resuppressed(tmp_path):
    """v3 structural fingerprint: a defect that disappears and reappears in a genuinely
    different shape is NOT silently re-suppressed (external review). Same shape/position
    stays suppressed."""
    p = new_prs()
    box = tb(p.slides.add_slide(p.slide_layouts[6]) if False else add_slide(p),
             1, 1, 5, 0.5, "모노 폴백 한글", font="IBM Plex Mono", size=12)
    box.name = "KPI title"
    deck = save(p, tmp_path, "b.pptx")
    bl = os.path.join(str(tmp_path), "bl.json")
    run_cli([deck, "--write-baseline", bl])
    # same shape name + position: suppressed
    doc = json.loads(run_cli([deck, "--json", "--baseline", bl]).stdout)
    assert doc["summary"]["baseline_suppressed"] == 1 and doc["summary"]["error_count"] == 0
    # same defect on a differently-named shape far away: a new finding, not suppressed
    p2 = new_prs()
    b2 = tb(add_slide(p2), 9, 5, 3, 0.5, "모노 폴백 한글", font="IBM Plex Mono", size=12)
    b2.name = "Footnote source"
    moved = save(p2, tmp_path, "moved.pptx")
    doc = json.loads(run_cli([moved, "--json", "--baseline", bl]).stdout)
    assert doc["summary"]["error_count"] == 1, doc["summary"]

def test_baseline_v3_generic_name_uses_position(tmp_path):
    """0.7.1 P0: with the generator default names ('TextBox N'), a defect that moves to a
    far-away box must NOT be re-suppressed. Stripping the counter used to collapse both
    to 'textbox' and hide the moved finding."""
    p = new_prs()
    b1 = tb(add_slide(p), 1, 1, 4, 0.5, "모노 폴백 한글", font="IBM Plex Mono", size=12)
    b1.name = "TextBox 3"   # generator default
    deck = save(p, tmp_path, "b.pptx")
    bl = os.path.join(str(tmp_path), "bl.json")
    run_cli([deck, "--write-baseline", bl])
    # same generic name, far-away position: identity is the bbox, so it is a NEW finding
    p2 = new_prs()
    b2 = tb(add_slide(p2), 9, 6, 3, 0.5, "모노 폴백 한글", font="IBM Plex Mono", size=12)
    b2.name = "TextBox 7"
    moved = save(p2, tmp_path, "moved.pptx")
    doc = json.loads(run_cli([moved, "--json", "--baseline", bl]).stdout)
    assert doc["summary"]["error_count"] == 1, doc["summary"]
    # same generic name, same position: still suppressed
    same = json.loads(run_cli([deck, "--json", "--baseline", bl]).stdout)
    assert same["summary"]["error_count"] == 0 and same["summary"]["baseline_suppressed"] == 1

def test_severity_override_policy_rules_only(tmp_path):
    """0.8.x (external audit): per-rule severity override, restricted to the policy
    layer. E2 demoted to warning passes the default gate but stays visible (JSON
    severity, summary.severity_overrides); a mechanical gate (E1) cannot be demoted;
    W14 can be turned off."""
    d = os.path.join(str(tmp_path), "deck")
    os.makedirs(d)
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 8, 0.6, "구조적" + EM_DASH + "개선", size=14, ea="맑은 고딕")  # E2 only
    deck = save(p, d, "styled.pptx")
    with open(os.path.join(d, ".archforge.json"), "w", encoding="utf-8") as f:
        f.write('{"profile": "full", "severity": {"E2": "warning"}}')
    r = run_cli([deck, "--json"])
    assert r.returncode == 0, (r.returncode, r.stdout)   # demoted: no longer blocks
    doc = json.loads(r.stdout)
    assert doc["summary"]["error_count"] == 0 and doc["summary"]["warn_count"] >= 1
    assert doc["summary"]["severity_overrides"] == {"E2": "warning"}
    assert any(w["code"] == "E2" for w in doc["warnings"])
    # schema 2: the finding's severity says warning (list membership, not the registry)
    v2 = json.loads(run_cli([deck, "--schema", "2", "--json"]).stdout)
    e2 = [f for f in v2["findings"] if f["code"] == "E2"][0]
    assert e2["severity"] == "warning"
    # JUnit agrees: E2 is not a <failure> when demoted
    import xml.etree.ElementTree as ET
    jx = os.path.join(str(tmp_path), "j.xml")
    run_cli([deck, "--junit", jx])
    e2case = [c for c in ET.parse(jx).getroot().iter("testcase")
              if c.get("name").startswith("E2")][0]
    assert e2case.find("failure") is None
    # mechanical gates cannot be demoted
    with open(os.path.join(d, ".archforge.json"), "w", encoding="utf-8") as f:
        f.write('{"severity": {"E1": "warning"}}')
    r = run_cli([deck])
    assert r.returncode == 2 and "mechanical" in (r.stderr or "")
    # off drops the finding entirely
    with open(os.path.join(d, ".archforge.json"), "w", encoding="utf-8") as f:
        f.write('{"profile": "full", "severity": {"E2": "off"}}')
    doc = json.loads(run_cli([deck, "--json"]).stdout)
    assert not any(x["code"] == "E2" for x in doc["errors"] + doc["warnings"])

def test_baseline_artifact_identity_and_inspect(tmp_path):
    """0.8: a baseline records the deck it was written from; applying it to a
    differently-named deck warns, and `baseline inspect` shows what is suppressed."""
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "모노 폴백 한글", font="IBM Plex Mono", size=12)
    deck = save(p, tmp_path, "quarterly.pptx")
    bl = os.path.join(str(tmp_path), "bl.json")
    run_cli([deck, "--write-baseline", bl])
    with open(bl, encoding="utf-8") as f:
        doc = json.load(f)
    assert doc["artifact"]["file_name"] == "quarterly.pptx"
    assert len(doc["artifact"]["sha256_12"]) == 12
    # same deck under a different name: still applies (warning, baselines are beta),
    # and the mismatch is surfaced on stderr
    other = save(p, tmp_path, "unrelated.pptx")
    r = run_cli([other, "--baseline", bl])
    assert "quarterly.pptx" in (r.stderr or "")
    # inspect: machine-readable summary of what the baseline suppresses
    r = run_cli(["baseline", "inspect", bl, "--json"])
    assert r.returncode == 0
    info = json.loads(r.stdout)
    assert info["schema_version"] == "3"
    assert info["suppressed_by_code"].get("E1") == 1
    assert info["artifact"]["file_name"] == "quarterly.pptx"

def test_baseline_v2_rejected(tmp_path):
    """A v2 baseline is rejected with a regenerate message (the one migration ADR 004
    committed to)."""
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "모노 폴백 한글", font="IBM Plex Mono", size=12)
    deck = save(p, tmp_path, "b.pptx")
    bl = os.path.join(str(tmp_path), "v2.json")
    with open(bl, "w", encoding="utf-8") as f:
        f.write('{"schema_version": "2", "findings": [{"code": "E1", "fingerprint": "x", "count": 1}]}')
    r = run_cli([deck, "--baseline", bl])
    assert r.returncode == 2 and "regenerate" in (r.stderr or "").lower()
