# -*- coding: utf-8 -*-
"""Finding 모델(0.4.0, 3차 외부 리뷰 구조 개편).

검사 엔진은 로케일 중립 finding(코드 + 메시지 id + 포맷 인자)만 만들고, 사람 언어는
리포터 단계(message 프로퍼티 접근 시점)에서 결정된다. 같은 결과를 여러 언어로 재출력할
수 있고, 메시지 문구 변경이 판정 스냅샷을 흔들지 않는다.

하위호환: 기존 소비자(deck_lint, test_gates, 사용자 코드)는 finding을
(page, code, message, detail) 4튜플로 언패킹·인덱싱한다. Finding은 그 계약을
그대로 지원한다(len 4, 순회, 인덱스).
"""
from typing import Dict, Optional, Tuple

try:
    from .messages import M
except ImportError:   # 파일 단독 실행 폴백
    from messages import M


class Finding:
    """단일 검사 결과. msg_id+args가 정본이고 message는 현재 언어로 렌더된 뷰다.

    loc(있으면)는 에이전트 자동 수정용 구조화 타깃:
    {"shape_id", "shape_name", "paragraph", "run", "bbox"[x,y,w,h in], "part"} 중 가용 키.
    """

    __slots__ = ("page", "code", "msg_id", "args", "detail", "loc", "_msg_override")

    def __init__(self, page: int, code: str, msg_id: str, args: Tuple = (),
                 detail: str = "", loc: Optional[Dict] = None,
                 _msg_override: Optional[str] = None):
        self.page = page
        self.code = code
        self.msg_id = msg_id
        self.args = tuple(args)
        self.detail = detail
        self.loc = loc
        # e3의 autofit 주석처럼 인자 안에 이미 렌더된 조각이 필요한 극소수 케이스,
        # 그리고 구형 코드가 완성 문자열을 넘기는 이행기 경로용(정본은 msg_id).
        self._msg_override = _msg_override

    @property
    def message(self) -> str:
        if self._msg_override is not None:
            return self._msg_override
        tpl = M(self.msg_id)
        return tpl % self.args if self.args else tpl

    # ---- (page, code, message, detail) 4튜플 하위호환 계약
    def _as_tuple(self):
        return (self.page, self.code, self.message, self.detail)

    def __iter__(self):
        return iter(self._as_tuple())

    def __getitem__(self, i):
        return self._as_tuple()[i]

    def __len__(self):
        return 4

    def __repr__(self):
        return "Finding(p%02d %s %s)" % (self.page, self.code, self.msg_id)

    def to_dict(self) -> Dict:
        d = {"page": self.page, "code": self.code, "message": self.message,
             "detail": self.detail}
        if self.loc:
            d["location"] = self.loc
        return d

    def fingerprint(self) -> str:
        """baseline 대조용 안정 지문: 페이지+코드+detail(메시지 언어와 무관)."""
        import hashlib
        raw = "%d|%s|%s" % (self.page, self.code, self.detail)
        return hashlib.sha1(raw.encode("utf-8")).hexdigest()[:16]


def shape_loc(sp, paragraph: Optional[int] = None, run: Optional[int] = None,
              part: Optional[str] = None) -> Optional[Dict]:
    """도형에서 구조화 위치 추출(실패 항목은 생략, 전부 실패면 None)."""
    loc: Dict = {}
    try:
        loc["shape_id"] = sp.shape_id
    except Exception:
        pass
    try:
        if sp.name:
            loc["shape_name"] = sp.name
    except Exception:
        pass
    try:
        if None not in (sp.left, sp.top, sp.width, sp.height):
            emu = 914400.0
            loc["bbox"] = [round(sp.left / emu, 3), round(sp.top / emu, 3),
                           round(sp.width / emu, 3), round(sp.height / emu, 3)]
    except Exception:
        pass
    if paragraph is not None:
        loc["paragraph"] = paragraph
    if run is not None:
        loc["run"] = run
    if part:
        loc["part"] = part
    return loc or None
