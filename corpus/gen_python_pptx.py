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
    p.save(path)
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


if __name__ == "__main__":
    main()
