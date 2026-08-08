# -*- coding: utf-8 -*-
"""Document-order iteration over a paragraph's inline items: a:r, a:fld, a:br.

The first slice of the #5 document model, and the first one that is a unification rather
than a move. This walk existed twice, written independently: once in lint.py feeding the
E1-E4/W1/W5/W8 gates, once in detectors_geometry.py feeding the W15-W17 glyph boxes. Same
fallback logic, same count-mismatch escape hatch, free to diverge -- and paragraph text is
the one thing the typography and geometry sides must agree on, because E2 reads dash
context out of the same string the glyph boxes measure.

python-pptx's para.runs returns a:r elements only. A slide-number field (a:fld) renders
with the same rPr rules as a normal run and an a:br starts a new visual line, so a walker
that sees only a:r misses text that PowerPoint draws (0.5.0 and 0.6.0, both from review
findings that shipped). This module is why neither consumer has to know that again.

Pure parsing: lxml elements in, plain tuples out. No Finding, no I/O, no sibling imports
beyond the namespace constant.
"""
try:
    from .ooxml import NS
except ImportError:   # standalone execution
    from ooxml import NS


class _FldRun:
    """Adapter for a:fld (auto fields such as slide number, date). Since CT_TextField also has
    an rPr+t structure per the schema, PowerPoint renders it with the same rules as a normal
    run, but python-pptx's para.runs only returns a:r, so field text was a blind spot for
    E1/E3/E4 (carried over from the fourth review, 0.5.0). Exposes only the minimal interface
    the checking code uses: ._r for run_fonts/run_track, .font.size for size."""
    __slots__ = ("_r", "text")

    class _Pt:
        __slots__ = ("pt",)

        def __init__(self, pt):
            self.pt = pt

    def __init__(self, fld):
        self._r = fld
        t = fld.find(NS + "t")
        self.text = (t.text or "") if t is not None else ""

    @property
    def font(self):
        return self   # only the .size access is used

    @property
    def size(self):
        try:
            v = self._r.find(NS + "rPr").get("sz")
        except Exception:
            return None
        if not v:
            return None
        try:
            return self._Pt(int(v) / 100.0)
        except (TypeError, ValueError):
            return None


def iter_inline_items(para):
    """The paragraph's inline items in document order, as (run_like, run_index, is_fld).

    run_like is a python-pptx run for a:r, a _FldRun for a:fld, or None for a:br (one
    line-break character in E2 context and offsets; a new visual line in glyph geometry).
    run_index is the index into para.runs for a:r items and None otherwise.

    If the a:r count under the XML does not match para.runs, or the walk raises at all,
    falls back to plain para.runs order. Both call sites carried exactly this escape hatch
    independently; a structural surprise degrades to the pre-0.5.0 behaviour instead of
    guessing at document order.
    """
    runs = list(para.runs)
    items = []
    try:
        r_seen = 0
        for child in para._p:
            tag = child.tag
            if tag == NS + "r":
                if r_seen < len(runs):
                    items.append((runs[r_seen], r_seen, False))
                    r_seen += 1
            elif tag == NS + "br":
                items.append((None, None, False))
            elif tag == NS + "fld":
                items.append((_FldRun(child), None, True))
        if r_seen != len(runs):
            items = [(r, i, False) for i, r in enumerate(runs)]
    except Exception:
        items = [(r, i, False) for i, r in enumerate(runs)]
    return items
