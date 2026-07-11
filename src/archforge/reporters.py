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


def build_json_doc(path: str, errors: List, warns: List, ghost, summary: Dict,
                   schema: str = "1.0", capabilities: Optional[Dict] = None,
                   abstentions: Optional[List] = None,
                   invocation: Optional[Dict] = None,
                   rules_split: Optional[Dict] = None) -> Dict:
    """schema "1.0" (default, stable through 1.x): split errors[]/warnings[], no per-item
    severity. schema "2.0" (0.7, opt-in via --schema 2): a single findings[] array with
    severity and structured `data` on each item, plus `capabilities`, structured
    `abstentions[]`, an `invocation` block (profile/policy/config/thresholds), and a
    `rules` split (executed / profile_excluded / user_suppressed) so a consumer can tell
    why a rule produced nothing (0.7.1). Both shapes carry the same verdict."""
    if schema == "2.0":
        findings = ([f.to_dict("2.0") for f in errors] +
                    [f.to_dict("2.0") for f in warns])
        return {
            "schema_version": "2.0",
            "tool": {"name": "archforge", "version": _tool_version()},
            "target_renderer": "powerpoint-windows",
            "file": path,
            "lang": get_lang(),
            "invocation": invocation or {},
            "rules": rules_split or {},
            "findings": findings,
            "capabilities": capabilities or {},
            "abstentions": abstentions or [],
            "ghost": [{"page": si, "title": t} for si, t in (ghost or [])],
            "summary": summary,
        }
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


def _junit_sanitize(s):
    return "".join(ch for ch in s if ch == "\t" or ch == "\n" or ord(ch) >= 0x20)


def build_junit_multi(items: List) -> str:
    """JUnit XML for CI systems that consume test reports (Jenkins, GitLab, ...).

    items = [(path, errors, warns, excluded_codes, policy, usage_error)], one testsuite
    per file, where policy = {"fail_on_warning": bool, "fail_incomplete": bool}. The
    honest mapping (issue #2 design discussion): a testcase is a RULE THAT RAN, so
    "tests" means "checks executed". ERROR findings map to <failure>; WARN findings fail
    only under fail_on_warning and otherwise land in <system-out> (JUnit <skipped> means
    "not run", so it is reserved for rules excluded by the active profile/--skip). W18 is
    a <failure> when fail_incomplete is set, so the JUnit verdict matches the CLI exit
    code instead of a suite reading green while the run exited 1 (0.7.1, external review
    P0). A file that could not be checked at all becomes a single <error> testcase."""
    import xml.etree.ElementTree as ET
    from .rules import ALL_CODES, TITLES

    def _order(c):
        return (c[0] != "E", int(c[1:]))

    root = ET.Element("testsuites", name="archforge")
    total = failures = errors_n = skipped_n = 0
    for path, errs, warns, excluded, policy, usage_error in items:
        if isinstance(policy, bool):   # back-compat: a bare warn_fail flag
            policy = {"fail_on_warning": policy, "fail_incomplete": False}
        warn_fail = policy.get("fail_on_warning", False)
        fail_incomplete = policy.get("fail_incomplete", False)
        suite = ET.SubElement(root, "testsuite", name=path)
        if usage_error is not None:
            case = ET.SubElement(suite, "testcase", classname=path, name="usage")
            err = ET.SubElement(case, "error", message=_junit_sanitize(usage_error))
            err.text = _junit_sanitize(usage_error)
            suite.set("tests", "1")
            suite.set("errors", "1")
            suite.set("failures", "0")
            suite.set("skipped", "0")
            total += 1
            errors_n += 1
            continue
        by_code: Dict[str, List] = {}
        for f in list(errs) + list(warns):
            by_code.setdefault(f.code, []).append(f)
        s_tests = s_fail = s_skip = 0
        for code in sorted(ALL_CODES, key=_order):
            case = ET.SubElement(suite, "testcase", classname=path,
                                 name="%s %s" % (code, TITLES.get(code, code)))
            s_tests += 1
            if code in excluded:
                ET.SubElement(case, "skipped",
                              message="excluded by the active profile or --skip")
                s_skip += 1
                continue
            hits = by_code.get(code, [])
            if not hits:
                continue
            lines = "\n".join(_junit_sanitize("p%02d %s | %s" % (f.page, f.message, f.detail))
                              for f in hits)
            is_fail = code.startswith("E") or warn_fail or \
                (code == "W18" and fail_incomplete)
            if is_fail:
                fail = ET.SubElement(case, "failure",
                                     message="%d finding(s)" % len(hits))
                fail.text = lines
                s_fail += 1
            else:
                out = ET.SubElement(case, "system-out")
                out.text = lines
        suite.set("tests", str(s_tests))
        suite.set("failures", str(s_fail))
        suite.set("errors", "0")
        suite.set("skipped", str(s_skip))
        total += s_tests
        failures += s_fail
        skipped_n += s_skip
    root.set("tests", str(total))
    root.set("failures", str(failures))
    root.set("errors", str(errors_n))
    root.set("skipped", str(skipped_n))
    return '<?xml version="1.0" encoding="utf-8"?>\n' + ET.tostring(root, encoding="unicode")


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
                # The cross-run identity matches the baseline v3 fingerprint (structural,
                # location-aware) so GitHub code scanning and the baseline agree on when a
                # finding is "the same" instead of diverging (0.7.1, external review). The
                # v2 content-only fingerprint stays alongside for one release for tools
                # that pinned it.
                "partialFingerprints": {"archforgeFinding/v3": f.structural_fp(),
                                        "archforgeFinding/v2": f.fingerprint()},
            }
            props = {}
            if f.loc:
                props["target"] = f.loc
            data = f.data()
            if data:
                props["data"] = data
            if props:
                res["properties"] = props
            # Pair findings (W15 overlap, W17 straddle) carry the counterpart in loc.related;
            # surface it via the SARIF-standard relatedLocations, not only properties
            # (0.7.1, external review, section 10).
            rel = (f.loc or {}).get("related")
            if isinstance(rel, dict) and rel:
                phys = {"artifactLocation": {"uri": path.replace("\\", "/")}}
                bbox = rel.get("bbox")
                res["relatedLocations"] = [{
                    "physicalLocation": phys,
                    "message": {"text": "related shape %s" % (rel.get("shape_name")
                                                              or rel.get("shape_id") or "")},
                    "logicalLocations": [{"name": "slide %d" % f.page, "kind": "module"}],
                }]
                if bbox:
                    res["relatedLocations"][0]["properties"] = {"bbox": bbox}
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
