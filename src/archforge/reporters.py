# -*- coding: utf-8 -*-
"""Reporters (0.4.0): render a findings list as text/JSON/SARIF. Language is decided here
(findings are locale-neutral: the check/presentation split from the third external
review)."""
from typing import Dict, List, Optional

try:
    from .messages import M, get_lang
    from .rules import RULES, TITLES, severity
except ImportError:
    from messages import M, get_lang
    from rules import RULES, TITLES, severity


def _tool_version() -> str:
    try:
        from importlib.metadata import version
        return version("archforge")
    except Exception:
        return "unknown"


def build_json_doc(path: str, errors: List, warns: List, ghost, summary: Dict) -> Dict:
    return {
        "schema_version": "1.0",
        "tool": {"name": "archforge", "version": _tool_version()},
        "target_renderer": "powerpoint-windows",
        "file": path,
        "lang": get_lang(),
        "errors": [f.to_dict() for f in errors],
        "warnings": [f.to_dict() for f in warns],
        "ghost": [{"page": si, "title": t} for si, t in (ghost or [])],
        "summary": summary,
    }


def render_text(path: str, errors: List, warns: List, ghost,
                profile: str, profile_excl, skip,
                config_path: Optional[str] = None,
                baseline_suppressed: int = 0,
                baseline_path: Optional[str] = None) -> List[str]:
    lines = ["=== ARCHFORGE LINT: %s ===" % path]
    if ghost:
        lines.append(M("ghost_header"))
        for si, txt in ghost:
            lines.append("  p%02d  %s" % (si, txt[:60]))
    if not errors and not warns:
        lines.append("clean: ERROR 0, WARN 0")
    for f in errors:
        lines.append("  ERROR p%02d [%s] %s | %s" % (f.page, f.code, f.message, f.detail))
    for f in warns:
        lines.append("  WARN  p%02d [%s] %s | %s" % (f.page, f.code, f.message, f.detail))
    if profile != "full":
        lines.append(M("profile_applied") % (profile, ",".join(sorted(profile_excl))))
    if skip:
        lines.append(M("skip_applied") % ",".join(sorted(skip)))
    if config_path:
        # Trust-boundary visibility (fourth review): always show which config file
        # adjusted the gate
        lines.append(M("config_applied") % config_path)
    if baseline_suppressed:
        # Fixes baseline suppression being entirely invisible in human-facing output and
        # misread as 'clean'
        lines.append(M("baseline_applied") % (baseline_suppressed, baseline_path or ""))
    lines.append("--- ERROR %d, WARN %d ---" % (len(errors), len(warns)))
    return lines


def build_sarif(path: str, errors: List, warns: List) -> Dict:
    """Minimal valid SARIF 2.1.0 document (a shape GitHub code scanning accepts)."""
    return build_sarif_multi([(path, errors, warns)])


def build_sarif_multi(items: List) -> Dict:
    """items = [(path, errors, warns), ...]. When there are multiple files, merges them
    into one run with a per-file artifactLocation (scan mode, 0.5.0)."""
    rules_meta = []
    used = sorted({f.code for (_p, errs, ws) in items for f in list(errs) + list(ws)})
    for code in used:
        sev, cat, _msg_id = RULES.get(code, ("warning", "unknown", None))
        # Static titles, not parameterized finding templates: a rule shortDescription
        # with raw %.1fpt placeholders leaked into code-scanning UIs (0.6.0, external
        # review)
        rules_meta.append({
            "id": code,
            "shortDescription": {"text": TITLES.get(code, code)},
            "helpUri": "https://github.com/Love-Ash/archforge/blob/main/docs/rules/%s.md" % code,
            "defaultConfiguration": {"level": "error" if sev == "error" else "warning"},
            "properties": {"category": cat},
        })
    results = []
    for path, errors, warns in items:
        for f in list(errors) + list(warns):
            res = {
                "ruleId": f.code,
                "level": "error" if severity(f.code) == "error" else "warning",
                "message": {"text": "%s | %s" % (f.message, f.detail) if f.detail else f.message},
                "locations": [{
                    "physicalLocation": {
                        "artifactLocation": {"uri": path.replace("\\", "/")},
                    },
                    "logicalLocations": [{"name": "slide %d" % f.page, "kind": "module"}],
                }],
                # The baseline fingerprint doubles as a cross-run identity so GitHub code
                # scanning can track a finding between runs instead of re-opening it on
                # every push (0.6.0, external review)
                "partialFingerprints": {"archforgeFinding/v2": f.fingerprint()},
            }
            if f.loc:
                res["properties"] = {"target": f.loc}
            results.append(res)
    return {
        "$schema": "https://raw.githubusercontent.com/oasis-tcs/sarif-spec/master/Schemata/sarif-schema-2.1.0.json",
        "version": "2.1.0",
        "runs": [{
            "tool": {"driver": {"name": "archforge", "version": _tool_version(),
                                "informationUri": "https://github.com/Love-Ash/archforge",
                                "rules": rules_meta}},
            "results": results,
        }],
    }
