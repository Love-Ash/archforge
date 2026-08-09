# -*- coding: utf-8 -*-
"""OOXML primitives: the namespace prefixes and the EMU unit, and nothing else.

The lowest layer of the 0.7 decomposition (#5). It exists so the modules above it stop
each carrying their own copy of the same namespace string: `fonts.py` had defined an `NS`
identical to `lint.py`'s, and a third copy was one extraction away.

No parsing, no resolution, no findings. Anything that needs an lxml call belongs a layer up.
"""

# English Metric Units per inch. OOXML stores every coordinate and size in EMU.
EMU_PER_IN = 914400

# DrawingML: shapes, text runs, fills, effects, theme.
NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"

# PresentationML: slides, layouts, masters, placeholders.
NS_P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"


def canvas_size_in(path):
    """(width_in, height_in) of the slide canvas, read straight from
    ppt/presentation.xml's sldSz without loading the package through python-pptx.
    None on any failure -- the caller treats the canvas as unknown, never guesses.

    Exists for the JSON report: location.bbox is stated to the hundredth of an inch,
    but a consumer planning a W16 fix could not compute "move it back on-canvas" from
    the payload alone because the report never said how big the canvas is (#14)."""
    import re
    import zipfile
    try:
        blob = zipfile.ZipFile(path).read("ppt/presentation.xml").decode("utf-8", "ignore")
        m = re.search(r'<p:sldSz[^>]*cx="(\d+)"[^>]*cy="(\d+)"', blob)
        if not m:
            return None
        return (round(int(m.group(1)) / EMU_PER_IN, 3), round(int(m.group(2)) / EMU_PER_IN, 3))
    except Exception:
        return None
