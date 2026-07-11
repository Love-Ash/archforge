# -*- coding: utf-8 -*-
"""Config file and baseline (0.4.0, third external review structural overhaul).

Config discovery: --config explicit > deck file's folder > current folder, for
.archforge.json / .archforge.yml(.yaml). JSON is always supported with no dependency;
YAML only when PyYAML is present (`pip install archforge[yaml]`). CLI flags override the
config file.

Supported keys:
  profile: core|full|editorial
  lang: ko|en
  skip: [W14, ...]
  hard_min / body_min / small_min / w6_sim / w6_cluster: number
  baseline: baseline file path (recorded existing violations are suppressed; only new ones
    are reported)

The baseline file (schema 3) has the shape {"findings": [{"code","fingerprint","count"}...]}
plus run-condition metadata (tool_version/profile/lang/threshold_hash), produced by
`archforge deck.pptx --write-baseline PATH`. The v3 fingerprint is the locale-neutral
content key (page-independent, so it survives message-language changes and slide insertion)
combined with a structural location bucket, so a defect that moves to a genuinely different
place is not silently re-suppressed. v1/v2 files are rejected with a regenerate message.
"""
import json
import math
import os
from typing import Dict, List, Optional, Tuple

CONFIG_NAMES = (".archforge.json", ".archforge.yml", ".archforge.yaml")
_ALLOWED_KEYS = {"profile", "lang", "skip", "hard_min", "body_min", "small_min",
                 "w6_sim", "w6_cluster", "baseline", "severity"}


def find_config(deck_path: str, explicit: Optional[str] = None) -> Optional[str]:
    if explicit:
        return explicit if os.path.exists(explicit) else None
    dirs = []
    try:
        dirs.append(os.path.dirname(os.path.abspath(deck_path)))
    except Exception:
        pass
    dirs.append(os.getcwd())
    for d in dirs:
        for name in CONFIG_NAMES:
            p = os.path.join(d, name)
            if os.path.exists(p):
                return p
    return None


def load_config(path: str) -> Tuple[Dict, List[str]]:
    """(config dict, warning list). A quality gate's config must be fail-safe (fourth
    review): unknown keys (the incident where a typo'd profle=full silently ran as the
    default core) and type/range violations are errors, not something to ignore. The
    baseline path is resolved relative to the config file, not the execution directory."""
    warnings: List[str] = []
    if path.endswith((".yml", ".yaml")):
        try:
            import yaml   # optional extra: archforge[yaml]
        except ImportError:
            raise RuntimeError(
                "YAML config needs PyYAML: pip install archforge[yaml] "
                "(or use .archforge.json)")
        with open(path, encoding="utf-8") as f:
            data = yaml.safe_load(f) or {}
    else:
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
    if not isinstance(data, dict):
        raise RuntimeError("config root must be a mapping: %s" % path)
    unknown = sorted(k for k in data if k not in _ALLOWED_KEYS)
    if unknown:
        raise RuntimeError("unknown config key(s): %s (allowed: %s)"
                           % (", ".join(unknown), ", ".join(sorted(_ALLOWED_KEYS))))
    out = dict(data)
    # Type/range validation (a tidy error instead of a traceback)
    def _num(key, lo=None, lo_incl=False, hi=None):
        if key not in out:
            return
        # bool is an int subclass in Python, so float(True) == 1.0 silently passed as a
        # threshold (0.6.1, external review): reject it as a type error
        if isinstance(out[key], bool):
            raise RuntimeError("config %r must be a number, got %r" % (key, out[key]))
        try:
            v = float(out[key])
        except (TypeError, ValueError):
            raise RuntimeError("config %r must be a number, got %r" % (key, out[key]))
        # NaN passes every ordinary comparison below (NaN <= 0 is False) and json.load
        # accepts a bare NaN literal, which would silently disable gates like E3: the
        # exact failure mode this validation exists to block (0.6.0, external finding)
        if not math.isfinite(v):
            raise RuntimeError("config %r must be finite, got %r" % (key, out[key]))
        if lo is not None and (v <= lo if not lo_incl else v < lo):
            raise RuntimeError("config %r out of range: %r" % (key, out[key]))
        if hi is not None and v > hi:
            raise RuntimeError("config %r out of range: %r" % (key, out[key]))
        out[key] = v
    _num("hard_min", lo=0)
    _num("body_min", lo=0)
    _num("small_min", lo=0)
    _num("w6_sim", lo=0, hi=1)
    _num("w6_cluster", lo=1, lo_incl=True)
    if "w6_cluster" in out:
        # 1.9 silently truncating to 1 changed the gate without a trace (0.6.1):
        # only integral values are accepted
        if float(out["w6_cluster"]) != int(out["w6_cluster"]):
            raise RuntimeError("config 'w6_cluster' must be an integer, got %r"
                               % out["w6_cluster"])
        out["w6_cluster"] = int(out["w6_cluster"])
    if "profile" in out and out["profile"] not in ("full", "core", "editorial"):
        raise RuntimeError("config 'profile' must be full|core|editorial, got %r" % out["profile"])
    if "lang" in out and out["lang"] not in ("ko", "en"):
        raise RuntimeError("config 'lang' must be ko|en, got %r" % out["lang"])
    if "skip" in out and not (isinstance(out["skip"], list)
                              and all(isinstance(c, str) for c in out["skip"])):
        raise RuntimeError("config 'skip' must be a list of code strings")
    if "baseline" in out:
        if not isinstance(out["baseline"], str):
            raise RuntimeError("config 'baseline' must be a path string")
        out["baseline"] = os.path.normpath(
            os.path.join(os.path.dirname(os.path.abspath(path)), out["baseline"]))
    if "severity" in out:
        # Per-rule severity override (0.8.x, external audit: dash punctuation is not the
        # same class of defect as font fallback, and a house style that legitimately
        # uses dashes needs a knob short of --skip). Deliberately restricted to the
        # policy layer (the rules core excludes: E2, W6, W9-W14): the mechanical gates
        # (E1/E3/E4, size/geometry) and W18 are the product's trust anchor and cannot
        # be demoted by a deck-side file.
        try:
            from .rules import PROFILES
        except ImportError:
            from rules import PROFILES
        overridable = frozenset(PROFILES["core"])
        sev = out["severity"]
        if not isinstance(sev, dict):
            raise RuntimeError("config 'severity' must be an object of code -> level")
        norm = {}
        for code, level in sev.items():
            c = str(code).strip().upper()
            if c not in overridable:
                raise RuntimeError(
                    "config 'severity' may only override the policy-layer rules (%s); "
                    "%r is a mechanical gate" % (",".join(sorted(overridable)), c))
            if level not in ("error", "warning", "off"):
                raise RuntimeError(
                    "config 'severity' values must be error|warning|off, got %r" % (level,))
            norm[c] = level
        out["severity"] = norm
    return out, warnings


def deck_artifact(deck_path: str) -> Optional[Dict]:
    """Artifact identity for a baseline (0.8, external review: a baseline applied to an
    unrelated deck would still suppress shared fingerprints). Deliberately coarse: the
    file basename survives regeneration while a content hash would not, so the basename
    is the identity signal and the hash is recorded for exact-copy detection only."""
    try:
        import hashlib
        with open(deck_path, "rb") as f:
            digest = hashlib.sha256(f.read()).hexdigest()[:12]
        return {"file_name": os.path.basename(deck_path), "sha256_12": digest}
    except Exception:
        return None


def write_baseline(path: str, findings, profile: str = "", lang: str = "",
                   thresholds: Optional[Dict] = None,
                   artifact: Optional[Dict] = None) -> int:
    """Fingerprint v3 (schema 3): the v2 page-free content key plus a structural location
    bucket, so a defect that moves to a genuinely different place is treated as new rather
    than silently re-suppressed (external review). Records policy identity (profile,
    lang, tool version, threshold hash) and, when given, the artifact identity that the
    reader checks on load."""
    from collections import Counter
    counts = Counter()
    codes = {}
    total = 0
    for f in findings:
        fp = f.structural_fp()
        counts[fp] += 1
        codes[fp] = f.code
        total += 1
    try:
        from importlib.metadata import version
        tool_ver = version("archforge")
    except Exception:
        tool_ver = "unknown"
    thr_hash = ""
    if thresholds:
        import hashlib
        payload = ",".join("%s=%r" % (k, thresholds[k]) for k in sorted(thresholds))
        thr_hash = hashlib.sha1(payload.encode("utf-8")).hexdigest()[:12]
    doc = {
        "schema_version": "3",
        "tool_version": tool_ver,
        "profile": profile,
        "lang": lang,
        "threshold_hash": thr_hash,
        "findings": [{"code": codes[fp], "fingerprint": fp, "count": counts[fp]}
                     for fp in sorted(counts)],
    }
    if artifact:
        doc["artifact"] = artifact
    with open(path, "w", encoding="utf-8", newline="\n") as f:
        json.dump(doc, f, ensure_ascii=False, indent=2)
    return total


def load_baseline_meta(path: str) -> Dict:
    """The baseline's recorded run conditions (tool_version/profile/lang/threshold_hash).
    Callers compare these against the current run and warn on mismatch (0.6.0:
    recorded-only metadata was an audit comment, not a check)."""
    try:
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
        return {k: data.get(k) for k in ("tool_version", "profile", "lang",
                                         "threshold_hash", "artifact")}
    except Exception:
        return {}


def load_baseline(path: str) -> Dict[str, int]:
    """Structural fingerprint -> allowed occurrence count. v1/v2 files require
    regeneration (the fingerprint now includes a location bucket)."""
    with open(path, encoding="utf-8") as f:
        data = json.load(f)
    if str(data.get("schema_version")) != "3":
        raise RuntimeError(
            "baseline schema %r is outdated; regenerate with --write-baseline"
            % data.get("schema_version"))
    out: Dict[str, int] = {}
    for e in data.get("findings", []):
        out[e["fingerprint"]] = out.get(e["fingerprint"], 0) + int(e.get("count", 1))
    return out


def apply_baseline(findings, known: Dict[str, int]):
    """(new findings list, suppressed count). Suppresses only up to the allowed occurrence
    count per structural fingerprint (multiset semantics). W18 is not eligible for baseline
    suppression (it signals incompleteness)."""
    budget = dict(known)
    kept, suppressed = [], 0
    for f in findings:
        if f.code != "W18":
            fp = f.structural_fp()
            if budget.get(fp, 0) > 0:
                budget[fp] -= 1
                suppressed += 1
                continue
        kept.append(f)
    return kept, suppressed
