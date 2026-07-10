# -*- coding: utf-8 -*-
"""리포터(0.4.0): finding 목록을 텍스트/JSON/SARIF로. 언어는 여기서 결정된다
(finding은 로케일 중립: 3차 외부 리뷰의 검사·표현 분리)."""
from typing import Dict, List, Optional

try:
    from .messages import M, get_lang
    from .rules import RULES, severity
except ImportError:
    from messages import M, get_lang
    from rules import RULES, severity


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
        # 신뢰 경계 가시성(4차 리뷰): 어떤 설정 파일이 게이트를 조정했는지 항상 표시
        lines.append(M("config_applied") % config_path)
    if baseline_suppressed:
        # baseline 억제가 사람용 출력에서 완전 불가시라 'clean'으로 오독되던 것 교정
        lines.append(M("baseline_applied") % (baseline_suppressed, baseline_path or ""))
    lines.append("--- ERROR %d, WARN %d ---" % (len(errors), len(warns)))
    return lines


def build_sarif(path: str, errors: List, warns: List) -> Dict:
    """SARIF 2.1.0 최소 유효 문서(GitHub code scanning 수용 형태)."""
    rules_meta = []
    used = sorted({f.code for f in list(errors) + list(warns)})
    for code in used:
        sev, cat, msg_id = RULES.get(code, ("warning", "unknown", None))
        rules_meta.append({
            "id": code,
            "shortDescription": {"text": M(msg_id) if msg_id else code},
            "properties": {"category": cat},
        })
    results = []
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
