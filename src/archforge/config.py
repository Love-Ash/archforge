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
    """(설정 dict, 경고 목록). 알 수 없는 키는 경고로 알리고 무시한다."""
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
    out = {}
    for k, v in data.items():
        if k not in _ALLOWED_KEYS:
            warnings.append("unknown config key ignored: %s" % k)
            continue
        out[k] = v
    return out, warnings


def write_baseline(path: str, findings) -> int:
    entries = [{"code": f.code, "page": f.page, "fingerprint": f.fingerprint()}
               for f in findings]
    with open(path, "w", encoding="utf-8", newline="\n") as f:
        json.dump({"schema_version": "1.0", "findings": entries}, f,
                  ensure_ascii=False, indent=2)
    return len(entries)


def load_baseline(path: str) -> set:
    with open(path, encoding="utf-8") as f:
        data = json.load(f)
    return {e["fingerprint"] for e in data.get("findings", [])}


def apply_baseline(findings, known: set):
    """(신규 finding 목록, 억제 수). W18은 baseline 대상이 아니다(불완전성 신호)."""
    kept, suppressed = [], 0
    for f in findings:
        if f.code != "W18" and f.fingerprint() in known:
            suppressed += 1
            continue
        kept.append(f)
    return kept, suppressed
