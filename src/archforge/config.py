# -*- coding: utf-8 -*-
"""설정 파일과 baseline(0.4.0, 3차 외부 리뷰 구조 개편).

설정 탐색: --config 명시 > 덱 파일 폴더 > 현재 폴더에서 .archforge.json /
.archforge.yml(.yaml). JSON은 의존성 없이 항상 지원하고, YAML은 PyYAML이 있을 때만
(`pip install archforge[yaml]`). CLI 플래그가 설정 파일을 이긴다.

지원 키:
  profile: core|full|editorial
  lang: ko|en
  skip: [W14, ...]
  hard_min / body_min / small_min / w6_sim / w6_cluster: 숫자
  baseline: baseline 파일 경로(기록된 기존 위반은 억제하고 신규만 보고)

baseline 파일은 {"findings": [{"code","page","fingerprint"}...]} 형태로
`archforge deck.pptx --write-baseline PATH`가 만든다. 지문은 페이지+코드+detail
기반이라 메시지 언어와 무관하다.
"""
import json
import os
from typing import Dict, List, Optional, Tuple

CONFIG_NAMES = (".archforge.json", ".archforge.yml", ".archforge.yaml")
_ALLOWED_KEYS = {"profile", "lang", "skip", "hard_min", "body_min", "small_min",
                 "w6_sim", "w6_cluster", "baseline"}


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
    """(설정 dict, 경고 목록). 품질 게이트의 설정은 fail-safe여야 한다(4차 리뷰):
    알 수 없는 키(오타 profle=full이 조용히 기본 core로 실행되는 사고)와 타입·범위
    위반은 무시가 아니라 오류다. baseline 경로는 실행 위치가 아니라 설정 파일 기준."""
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
    # 타입·범위 검증(traceback 대신 정돈된 오류)
    def _num(key, lo=None, lo_incl=False, hi=None):
        if key not in out:
            return
        try:
            v = float(out[key])
        except (TypeError, ValueError):
            raise RuntimeError("config %r must be a number, got %r" % (key, out[key]))
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
    return out, warnings


def write_baseline(path: str, findings, profile: str = "", lang: str = "") -> int:
    """지문 v2(스키마 2): 페이지 무관 지문 + 발생 수(count) + 실행 조건 메타데이터.
    v1의 세 결함(언어 의존 detail, 페이지 삽입 취약, multiset 소실)의 교정(4차 리뷰)."""
    from collections import Counter
    counts = Counter()
    codes = {}
    total = 0
    for f in findings:
        fp = f.fingerprint()
        counts[fp] += 1
        codes[fp] = f.code
        total += 1
    try:
        from importlib.metadata import version
        tool_ver = version("archforge")
    except Exception:
        tool_ver = "unknown"
    doc = {
        "schema_version": "2",
        "tool_version": tool_ver,
        "profile": profile,
        "lang": lang,
        "findings": [{"code": codes[fp], "fingerprint": fp, "count": counts[fp]}
                     for fp in sorted(counts)],
    }
    with open(path, "w", encoding="utf-8", newline="\n") as f:
        json.dump(doc, f, ensure_ascii=False, indent=2)
    return total


def load_baseline(path: str) -> Dict[str, int]:
    """지문 -> 허용 발생 수. v1(스키마 1.0) 파일은 재생성을 요구한다(지문 체계 변경)."""
    with open(path, encoding="utf-8") as f:
        data = json.load(f)
    if str(data.get("schema_version")) != "2":
        raise RuntimeError(
            "baseline schema %r is outdated; regenerate with --write-baseline"
            % data.get("schema_version"))
    out: Dict[str, int] = {}
    for e in data.get("findings", []):
        out[e["fingerprint"]] = out.get(e["fingerprint"], 0) + int(e.get("count", 1))
    return out


def apply_baseline(findings, known: Dict[str, int]):
    """(신규 finding 목록, 억제 수). 지문당 허용 발생 수까지만 억제한다(multiset 의미).
    W18은 baseline 대상이 아니다(불완전성 신호)."""
    budget = dict(known)
    kept, suppressed = [], 0
    for f in findings:
        if f.code != "W18":
            fp = f.fingerprint()
            if budget.get(fp, 0) > 0:
                budget[fp] -= 1
                suppressed += 1
                continue
        kept.append(f)
    return kept, suppressed
