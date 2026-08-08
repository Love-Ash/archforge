# -*- coding: utf-8 -*-
"""Copy detectors: the two rules that read what a deck says rather than how it is drawn.

  W11  buzzword clusters and stock opening cliches
  W14  titles that name a topic instead of making a claim, in Hangul and in English

Extracted from lint.py for the 0.7 decomposition (#5). These emit Findings, so unlike the
kernels below them they are not pure, but they take text and return findings and touch no
file and no CLI. Both are policy rather than physics, which is why they sit behind the
full profile and why their word lists live here next to the code that reads them.

Re-exported from lint for backward compatibility.
"""
import re

try:
    from .findings import Finding
    from .scripts import is_cjk
except ImportError:   # standalone execution
    from findings import Finding
    from scripts import is_cjk

# AI-tell copy: only obvious cliches (a narrow dictionary = suppresses false positives).
# General words that could be legitimate in context are not included.
BUZZWORDS = (
    "시너지", "패러다임", "게임체인저", "게임 체인저", "혁신을 가속", "가치를 극대화",
    "미래를 선도", "새로운 지평", "무한한 가능성", "홀리스틱", "엔드투엔드", "엔드 투 엔드",
    "초격차", "글로벌 리더로 도약", "위대한 여정",
    "synergy", "paradigm shift", "game-changer", "game changer", "cutting-edge",
    "state-of-the-art", "seamless", "revolutionize", "leverage synerg", "holistic",
    "unlock the potential", "empower",
)

STALE_OPENINGS = (
    "오늘날", "급변하는", "4차 산업혁명 시대", "바야흐로", "현대 사회에서",
    "디지털 전환의 시대", "디지털 전환의 물결", "알아보겠습니다", "살펴보겠습니다",
    "in today's", "in the rapidly changing", "in this presentation",
)

def copy_cliche_check(page_texts, warns):
    """W11: AI-tell copy. Buzzwords are checked on every page, cliche openings only in the
    intro (p1-3)."""
    for si in sorted(page_texts):
        blob = " ".join(page_texts[si])
        low = blob.lower()
        hits = sorted({b for b in BUZZWORDS if b.lower() in low})
        if hits:
            warns.append(Finding(si, "W11", "w11_buzz", (len(hits),), ", ".join(hits[:5])))
        if si <= 3:
            op = sorted({o for o in STALE_OPENINGS if o.lower() in low})
            if op:
                warns.append(Finding(si, "W11", "w11_open", (), ", ".join(op[:5])))

# English finite-verb forms that mark a claim-style title. Prefer false negatives over
# false positives: copulas/auxiliaries and ambiguous noun/verb homographs are omitted.
# Measured against the issue fixtures (noun-phrase decks fire; verb/number+unit decks do not).
_EN_CLAIM_VERBS = frozenset({
    "grows", "grew", "rises", "rose", "fall", "falls", "fell", "jumps", "jumped",
    "drives", "drove", "reaches", "reached", "beats", "misses", "missed",
    "expands", "expanded", "cuts", "gains", "gained", "loses", "lost",
    "improves", "improved", "declines", "declined", "surges", "surged",
    "drops", "dropped", "hits", "tops", "topped", "leads", "led", "lags",
    "lagged", "remains", "remained", "continues", "continued", "accelerates",
    "accelerated", "slows", "slowed", "climbs", "climbed", "slides", "slid",
    "boosts", "boosted", "lifts", "lifted", "shrinks", "shrank", "narrows",
    "narrowed", "widens", "widened", "outpaces", "outpaced", "outperforms",
    "outperformed", "underperforms", "underperformed", "adds", "added",
    "opens", "opened", "closes", "closed", "launches", "launched", "ships",
    "shipped", "wins", "won", "fails", "failed", "holds", "held", "keeps",
    "kept", "sets", "set", "breaks", "broke", "posts", "posted", "reports",
    "reported", "delivers", "delivered", "fuels", "fueled", "fuelled",
})

# Structural English titles that should not enter the W14 eligibility pool.
# Cover/section/closing slides are noun phrases by nature; counting them would
# fire W14 on any three-slide deck with a title page and divider.
_EN_STRUCTURAL_TITLES = frozenset({
    "cover", "title", "title page", "agenda", "introduction",
    "overview", "appendix", "thank you", "thanks", "q&a", "q and a", "qa",
    "closing", "contents", "table of contents",
    "toc", "outline", "references", "bibliography", "the end", "end",
    "summary", "background", "agenda overview",
})

def _en_title_key(title: str) -> str:
    key = re.sub(r"\s+", " ", title.strip().lower())
    key = re.sub(r"[^a-z0-9 &]+", "", key).strip()
    return key

def _en_eligible_title(title: str, chars, latin_n: int, cjk_n: int) -> bool:
    """Whether an English title is substance enough for the W14 majority pool.

    Hangul eligibility is a content share test; English needs more than "is Latin"
    so cover/divider/closing slides do not flood the pool on clean decks.
    """
    if cjk_n >= 3:
        return False
    if latin_n < 4 or latin_n < 0.6 * len(chars):
        return False
    if _en_title_key(title) in _EN_STRUCTURAL_TITLES:
        return False
    words = re.findall(r"[A-Za-z]{2,}", title)
    # Need either multi-word substance or enough letters that it is not a label.
    if len(words) >= 2 and sum(len(w) for w in words) >= 10:
        return True
    if len(words) >= 3:
        return True
    return False

def _en_is_claim(title):
    """English W14 claim signal: finite verb from a measured allowlist, or number+unit.
    Bare noun phrases return False. Prefer false negatives on editorial headlines."""
    if "?" in title or "!" in title:
        return True
    # number + unit (%, pp/bp, x multiplier, currency); keeps the Korean contract's spirit
    if re.search(
        r"(?:[$€£]\s*[0-9]|[0-9][0-9,.]*\s*(?:%|pp|bp|[xX]|million|billion|thousand|percent))",
        title,
    ):
        return True
    words = re.findall(r"[A-Za-z]+", title)
    return any(w.lower() in _EN_CLAIM_VERBS for w in words)

# Cover / divider / closing slides in Korean decks. A structural slide is not part of the
# deck's argument, so counting it inflates W14 on both sides of the ratio. Mirrors
# _EN_STRUCTURAL_TITLES, which landed for English in #9.
_KO_STRUCTURAL_TITLES = frozenset({
    "표지", "목차", "차례", "부록", "별첨", "감사합니다", "감사", "질의응답", "질의 응답",
    "마무리", "맺음말", "끝", "참고문헌", "참고 문헌", "출처", "회사소개", "회사 소개",
    "들어가며", "시작하며", "목차 및 구성", "발표를 마칩니다", "이상입니다",
    "경청해 주셔서 감사합니다",
})

# Korean nouns whose last syllable is also a sentence ending. The ending test cannot tell
# 개요 (an overview) from 해요 (a predicate), so the collisions are listed rather than
# guessed at, the way the rule already carried a note about 바다.
#
# The list is incomplete on purpose and safe to leave that way: -자 is a productive agent
# suffix, so no hand-written list keeps up with it, and a word that is absent simply keeps
# the previous reading (counted as a claim, rule stays quiet). Every residual miss
# therefore lands on the false-negative side, which is the side W14 is supposed to err on.
# Words that can also be predicates in a title (가요, 포함, 약함, 이자) are deliberately
# left out: listing them would move errors to the firing side.
_KO_NOUN_FINALS = frozenset({
    "개요", "필요", "중요", "주요", "수요", "소요", "강요", "동요",
    "투자자", "사용자", "소비자", "참가자", "참여자", "가입자", "이용자", "구독자",
    "시청자", "개발자", "경영자", "창업자", "실무자", "담당자", "책임자", "관리자",
    "신청자", "응답자", "협력자", "수혜자", "지원자", "후보자", "방문자", "구매자",
    "판매자", "근로자", "노동자", "종사자", "대상자", "경쟁자", "설계자", "운영자",
    "제작자", "독자", "저자", "기자", "학자", "환자",
    "투자", "숫자", "글자", "문자", "의자", "모자", "상자", "전자", "한자",
    "게임", "모임", "책임", "쓰임",
    "바다", "소다",
})

# Sentence-final interrogatives written without a question mark. Recognising these moves a
# title out of the nominal count, which lowers the firing rate, so they are safe to add.
_KO_INTERROGATIVE_SUFFIX = ("인가", "는가", "은가", "던가", "을까")

_KO_SENTENCE_ENDINGS = ("다", "까", "요", "자", "죠", "함", "임")

_KO_NUMERIC_CLAIM = re.compile(
    r"[0-9][0-9,.]*\s*(%|배|억|조|만|천|pp|bp|x|X|원|건|명|개)")

def _ko_title_key(title):
    return re.sub(r"\s+", " ", title.strip()).rstrip(" ?!.…”’")

def _ko_is_claim(title):
    """Hangul W14 claim signal: a question, a number+unit, a sentence-final interrogative,
    or a sentence ending that is not the tail of a known noun.

    The number+unit branch is there because a title like "매출 3배 성장" is a claim-style
    headline even though it ends in a noun (external review, 2026-07-10)."""
    if "?" in title or "!" in title:
        return True
    if _KO_NUMERIC_CLAIM.search(title):
        return True
    core = _ko_title_key(title)
    if core.endswith(_KO_INTERROGATIVE_SUFFIX):
        return True
    if not core.endswith(_KO_SENTENCE_ENDINGS):
        return False
    return core.split(" ")[-1] not in _KO_NOUN_FINALS

def action_title_check(titles, warns):
    """W14: a majority of titles are descriptive noun phrases (e.g. "Market Overview,"
    "Competitive Analysis") = not action titles (the MBB idea that reading only the titles
    should carry the argument). Hangul titles use a sentence-ending heuristic minus the
    nouns that collide with it (see `_ko_is_claim`); English titles use a finite-verb
    allowlist or number+unit (see `_en_is_claim`). Structural titles are not eligible in
    either language (cover, agenda, appendix, closing / 표지, 목차, 부록, 감사합니다), so a
    clean deck with a title page and divider does not false-fire. Prefer false negatives
    on ambiguous verb/noun forms. Fires once per deck only when 3+ eligible titles are
    noun phrases and they make up at least half of eligible titles. Gate: full profile
    (editorial skips)."""
    entries = []
    for si in sorted(titles):
        txt = " ".join(titles[si][1]).strip()
        chars = [c for c in txt if not c.isspace()]
        if len(chars) < 4:
            continue
        cjk_n = sum(1 for c in chars if is_cjk(c))
        latin_n = sum(1 for c in chars if ("A" <= c <= "Z") or ("a" <= c <= "z"))
        # Hangul path: excludes titles with fewer than 3 Hangul characters or under 30%
        # Hangul share (big-stat numbers / short brand names; measured in the 50-deck scan).
        if cjk_n >= 3 and cjk_n >= 0.3 * len(chars):
            if _ko_title_key(txt) in _KO_STRUCTURAL_TITLES:
                continue
            entries.append((si, txt, _ko_is_claim(txt)))
            continue
        # English path: content-eligible Latin titles only (not structural covers).
        if _en_eligible_title(txt, chars, latin_n, cjk_n):
            entries.append((si, txt, _en_is_claim(txt)))
    nominal = [(si, t) for si, t, c in entries if not c]
    if len(nominal) >= 3 and len(nominal) * 2 >= len(entries):
        ex = " ".join("p%d'%s'" % (si, t[:14]) for si, t in nominal[:4])
        warns.append(Finding(0, "W14", "w14", (len(nominal), len(entries)), ex))
    return entries
