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
