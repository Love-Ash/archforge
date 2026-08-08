# -*- coding: utf-8 -*-
"""Seed corpus, python-pptx generator: one deck per flagship defect, constructed so the
ground truth is known by construction (the defect is deliberately seeded) rather than
by trusting the linter's own output. Expected findings live in the sibling .json
manifests consumed by corpus/run_corpus.py.

Usage: python corpus/gen_python_pptx.py
"""
import json
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "src"))

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.oxml.ns import qn

EM_DASH = chr(0x2014)
HERE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "python-pptx")


def _prs():
    p = Presentation()
    p.slide_width = Inches(13.333)
    p.slide_height = Inches(7.5)
    return p


def _tb(s, x, y, w, h, text, size=14, font=None, ea=None, spc=None):
    box = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    r = box.text_frame.paragraphs[0].add_run()
    r.text = text
    r.font.size = Pt(size)
    if font:
        r.font.name = font
    if ea:
        rPr = r._r.get_or_add_rPr()
        rPr.append(rPr.makeelement(qn("a:ea"), {"typeface": ea}))
    if spc is not None:
        r._r.get_or_add_rPr().set("spc", str(spc))
    return box


def emit(name, build, manifest):
    os.makedirs(HERE, exist_ok=True)
    path = os.path.join(HERE, name + ".pptx")
    p = _prs()
    build(p)
    # docProps carries python-pptx's bundled template properties, which name a
    # stranger and claim PowerPoint on a Mac wrote the file. Neither is true and both
    # ship in the sdist, so the writer is stated honestly instead. No rule reads
    # docProps, and the manifest's "generator" field is what records provenance.
    from archforge.demo import _save
    # docProps says "archforge corpus", not the library name. Provenance lives in the
    # manifest's "generator" field, which is where the corpus README points and where a
    # reader looks; putting the library name in docProps as well only ships a
    # generated-by tell in every sdist for information already recorded.
    _save(p, path, application="archforge corpus")
    manifest.setdefault("generator", "python-pptx")
    manifest.setdefault("profile", "full")
    manifest.setdefault("ground_truth", "by construction (defect deliberately seeded)")
    with open(os.path.join(HERE, name + ".json"), "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)
    print("wrote", path)


def main():
    def e1(p):
        s = p.slides.add_slide(p.slide_layouts[6])
        _tb(s, 1, 1, 6, 1, "아리알에 실린 한글", size=20, font="Arial")
    emit("e1_arial_hangul", e1, {"expected": {"E1": 1},
         "notes": "Hangul on a Latin-only a:latin with the default theme's empty ea slot"})

    def e2(p):
        s = p.slides.add_slide(p.slide_layouts[6])
        _tb(s, 1, 1, 8, 1, "growth was structural" + EM_DASH + "not cyclical", size=16,
            ea="맑은 고딕")
    emit("e2_em_dash", e2, {"expected": {"E2": 1},
         "notes": "word-to-word em dash; numeric ranges would pass"})

    def e3(p):
        s = p.slides.add_slide(p.slide_layouts[6])
        _tb(s, 1, 6.9, 5, 0.3, "source: internal accounts", size=4, ea="맑은 고딕")
    emit("e3_4pt", e3, {"expected": {"E3": 1}})

    def e4(p):
        s = p.slides.add_slide(p.slide_layouts[6])
        _tb(s, 1, 1, 6, 1, "자간이 벌어진 한글", size=16, spc=300, ea="맑은 고딕")
    emit("e4_tracked_hangul", e4, {"expected": {"E4": 1}})

    def w15(p):
        s = p.slides.add_slide(p.slide_layouts[6])
        _tb(s, 1.0, 2.4, 5.0, 1.0, "Revenue growth +18%", size=24, ea="맑은 고딕")
        _tb(s, 1.2, 2.5, 5.0, 1.0, "Operating margin 12.4%", size=24, ea="맑은 고딕")
    emit("w15_overlap", w15, {"expected": {"W15": 1}})

    def w16(p):
        s = p.slides.add_slide(p.slide_layouts[6])
        _tb(s, 12.0, 4.5, 3.0, 0.6, "Next quarter guidance", size=18, ea="맑은 고딕")
    emit("w16_offcanvas", w16, {"expected": {"W16": 1}})

    def clean(p):
        s = p.slides.add_slide(p.slide_layouts[6])
        _tb(s, 1, 1, 8, 1, "Quarterly results improved", size=20, ea="맑은 고딕")
        _tb(s, 1, 2.2, 8, 0.6, "구독 매출이 성장을 견인했습니다", size=14, ea="맑은 고딕")
    emit("clean_bilingual", clean, {"expected": {},
         "notes": "negative fixture: correct ea font, no dash, readable sizes"})

    EN_DASH = chr(0x2013)

    def e2_range_negative(p):
        s = p.slides.add_slide(p.slide_layouts[6])
        _tb(s, 1, 1, 8, 0.6, "FY2020" + EN_DASH + "2024 revenue up 18%", size=16,
            ea="맑은 고딕")
    emit("e2_range_negative", e2_range_negative, {"expected": {},
         "notes": "negative fixture: a numeric-range en dash is legitimate typography "
                  "and must pass (E2's exemption contract)"})

    def w1_small_body(p):
        s = p.slides.add_slide(p.slide_layouts[6])
        _tb(s, 1, 2, 9, 1.2,
            "This body-class paragraph runs well past forty characters of text "
            "so the frame is judged as body copy.", size=8, ea="맑은 고딕")
    emit("w1_small_body", w1_small_body, {"expected": {"W1": 1},
         "notes": "wide frame, long text, 8pt: below the 9pt body floor"})

    def w8_small_cjk(p):
        s = p.slides.add_slide(p.slide_layouts[6])
        _tb(s, 1, 1, 2.0, 0.4, "목업 안 라벨", size=6, ea="맑은 고딕")
    emit("w8_small_cjk", w8_small_cjk, {"expected": {"W8": 1},
         "notes": "narrow (<=4in) frame, 6pt CJK: the mockup-label pattern"})

    def w6_repeated_skeleton(p):
        for i in range(5):
            s = p.slides.add_slide(p.slide_layouts[6])
            _tb(s, 1, 0.8, 8, 0.6, "Section title %d" % i, size=24, ea="맑은 고딕")
            _tb(s, 1, 2.0, 6, 2.5, "Body block content", size=12, ea="맑은 고딕")
            _tb(s, 8, 2.0, 4, 2.5, "Side note block", size=12, ea="맑은 고딕")
    emit("w6_repeated_skeleton", w6_repeated_skeleton, {"expected": {"W6": 1, "W14": 1},
         "notes": "the same 3-frame skeleton on 5 pages: the recycled-grid tell; English noun-phrase section titles also trip W14 under the English path"})

    def w14_en_nominal(p):
        for t in ("Market Overview", "Competitive Analysis", "Product Lineup",
                  "Expansion Strategy", "Financial Plan", "Next Steps Summary"):
            s = p.slides.add_slide(p.slide_layouts[6])
            _tb(s, 1, 0.8, 9, 0.8, t, size=26, ea="맑은 고딕")
            _tb(s, 1, 2.2, 10, 3, "Supporting body copy", size=12, ea="맑은 고딕")
    emit("w14_en_nominal", w14_en_nominal, {"expected": {"W14": 1},
         "notes": "positive W14-EN fixture: English noun-phrase titles across pages"})

    def w14_en_action(p):
        for t in ("Market grows 18% on subscriptions", "Revenue reaches $12 million",
                  "Costs fall 8pp year over year", "Retention improves 2x",
                  "Margin expands after price cuts", "Pipeline adds 40% coverage"):
            s = p.slides.add_slide(p.slide_layouts[6])
            _tb(s, 1, 0.8, 9, 0.8, t, size=26, ea="맑은 고딕")
            _tb(s, 1, 2.2, 10, 3, "Supporting body copy", size=12, ea="맑은 고딕")
    emit("w14_en_action", w14_en_action, {"expected": {},
         "notes": "negative W14-EN fixture: finite verb or number+unit counts as a claim"})

    def _ko_deck(p, titles):
        for t in titles:
            s = p.slides.add_slide(p.slide_layouts[6])
            _tb(s, 1, 0.8, 9, 0.8, t, size=26, ea="맑은 고딕")
            _tb(s, 1, 2.2, 10, 3, "본문은 읽을 수 있는 크기로 들어갑니다.",
                size=12, ea="맑은 고딕")

    def w14_ko_nominal(p):
        # Every title here ends in a syllable that also ends a Korean sentence, so before
        # the noun-collision list they all read as claims and the rule stayed silent.
        _ko_deck(p, ("제품 개요", "사업 개요", "해외 투자", "핵심 사용자",
                     "모바일 게임", "시장 현황"))
    emit("w14_ko_nominal", w14_ko_nominal, {"expected": {"W14": 1},
         "notes": "positive W14 Hangul fixture: noun phrases whose final syllable collides "
                  "with a sentence ending (개요 / 투자 / 사용자 / 게임)"})

    def w14_ko_claim(p):
        # The mirror of the above: real predicates that end in the same syllables, so the
        # collision list may not swallow them.
        _ko_deck(p, ("매출이 전년 대비 늘었다", "성장세가 꺾였어요", "비용을 줄이자",
                     "점유율이 3분기에 반등했다", "이것이 최선인가",
                     "지금 진입해야 한다"))
    emit("w14_ko_claim", w14_ko_claim, {"expected": {},
         "notes": "negative W14 Hangul fixture: declarative, polite, propositive and "
                  "interrogative endings all count as claims. This one also passes on the "
                  "pre-2026-08-07 code, so it proves nothing about the noun-collision fix; "
                  "it is the regression guard against a future rewrite that enumerates "
                  "endings instead of nouns, which would read these as noun phrases and "
                  "fire on a deck of real claims"})

    def w14_ko_structural(p):
        # Three structural slides and two content noun phrases. Counting the structural
        # ones would put five nominal titles in the pool and fire; excluding them leaves
        # two, which is under the three-title minimum.
        _ko_deck(p, ("질의응답", "참고 문헌", "회사 소개", "시장 현황", "매출 추이"))
    emit("w14_ko_structural", w14_ko_structural, {"expected": {},
         "notes": "negative W14 Hangul fixture: cover/divider/closing titles are not part "
                  "of the deck argument and must not enter the majority pool"})


if __name__ == "__main__":
    main()
