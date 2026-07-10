# -*- coding: utf-8 -*-
"""Archforge: a Hangul-focused quality linter that machine-checks a built .pptx before it ships.

The name comes from arch (structure/architecture) and forge (to hammer and temper; a smithy).
It means a smithy that forges and refines a deck's structure and Hangul typography before it
ships.

`archforge.lint` is a submodule (it is not shadowed by a function of the same name on the
package). If a functional entry point is needed, use `from archforge.lint import lint` or
`archforge.lint_pptx`.
"""
from . import lint as _lint_module  # noqa: F401  (ensures the submodule is loaded)
from .lint import lint as lint_pptx  # noqa: F401
from .lint import frame_autofit, frame_font_scale, main  # noqa: F401
from .findings import Finding  # noqa: F401  (0.4.0 public model)

__version__ = "0.5.0"
