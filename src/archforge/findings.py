# -*- coding: utf-8 -*-
"""Finding model (0.4.0, third external review structural overhaul).

The check engine only produces locale-neutral findings (code + message id + format args); the
human-language rendering is decided at the reporter stage (when the message property is
accessed). The same result can be re-rendered in multiple languages, and message wording
changes do not shake the verdict snapshot.

Backward compatibility: existing consumers (deck_lint, test_gates, user code) unpack/index a
finding as a (page, code, message, detail) 4-tuple. Finding supports that contract as-is
(len 4, iteration, indexing).
"""
from typing import Dict, Optional, Tuple

try:
    from .messages import M
except ImportError:   # fallback for standalone file execution
    from messages import M


class Finding:
    """A single check result. msg_id+args is canonical and message is a view rendered in the
    current language.

    loc (if present) is a structured target for agent auto-fix: whichever keys are available
    among {"shape_id", "shape_name", "paragraph", "run", "bbox"[x,y,w,h in], "part"}.
    """

    __slots__ = ("page", "code", "msg_id", "args", "detail", "loc", "fp_key", "_msg_override")

    def __init__(self, page: int, code: str, msg_id: str, args: Tuple = (),
                 detail: str = "", loc: Optional[Dict] = None,
                 fp_key: Optional[str] = None,
                 _msg_override: Optional[str] = None):
        self.page = page
        self.code = code
        self.msg_id = msg_id
        self.args = tuple(args)
        self.detail = detail
        self.loc = loc
        # Locale-neutral key for baseline fingerprinting. For rules where detail mixes in a
        # translated string (W6, W10, W16), put only the data part here (fourth review
        # confirmed: this fixed a bug where a ko baseline was invalid under an en run).
        self.fp_key = fp_key
        # For the handful of cases that need an already-rendered fragment inside args, like
        # e3's autofit annotation, and for the transitional path where legacy code passes a
        # finished string (msg_id remains canonical).
        self._msg_override = _msg_override

    @property
    def message(self) -> str:
        if self._msg_override is not None:
            return self._msg_override
        tpl = M(self.msg_id)
        return tpl % self.args if self.args else tpl

    # ---- (page, code, message, detail) 4-tuple backward-compat contract
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
        """Stable fingerprint v2 for baseline comparison: code + locale-neutral content key.

        Deliberately excludes the page number: auto-generated decks commonly insert or
        reorder slides, so a page-based fingerprint became entirely invalid the moment a
        single slide was inserted (fourth review confirmed). When the same-content violation
        occurs in multiple places, the baseline manages it by occurrence count (count).
        Full identity for regenerated output is not possible without generator provenance
        (a source map), which is the 0.5+ roadmap."""
        import hashlib
        raw = "%s|%s" % (self.code, self.fp_key if self.fp_key is not None else self.detail)
        return hashlib.sha1(raw.encode("utf-8")).hexdigest()[:16]


def shape_loc(sp, paragraph: Optional[int] = None, run: Optional[int] = None,
              part: Optional[str] = None, cell: Optional[Tuple[int, int]] = None,
              xf: Optional[Tuple[float, float, float, float]] = None,
              bbox: Optional[list] = None, field: bool = False) -> Optional[Dict]:
    """Extract a structured location from a shape (failed items are omitted; None if
    everything fails).

    0.5.0 (carried over from the fourth review): a group child's raw left/top is in the
    group's chOff coordinate system, which was misaligned with slide coordinates. When given
    an xf=(ax,bx,ay,by) affine, this converts to absolute coordinates via abs = a*raw + b;
    when given a bbox argument, it uses the already-computed absolute bbox (in, [x,y,w,h])
    as-is (the effective glyph geometry for W15-W17). cell is the table cell (row, col),
    0-based; field=True marks an a:fld (auto field) origin (run index is based on para.runs,
    so fields have none)."""
    loc: Dict = {}
    if sp is not None:
        try:
            loc["shape_id"] = sp.shape_id
        except Exception:
            pass
        try:
            if sp.name:
                loc["shape_name"] = sp.name
        except Exception:
            pass
    if bbox is not None:
        try:
            loc["bbox"] = [round(float(v), 3) for v in bbox]
        except Exception:
            pass
    elif sp is not None:
        try:
            if None not in (sp.left, sp.top, sp.width, sp.height):
                emu = 914400.0
                ax, bx, ay, by = xf if xf is not None else (1.0, 0.0, 1.0, 0.0)
                loc["bbox"] = [round((ax * sp.left + bx) / emu, 3),
                               round((ay * sp.top + by) / emu, 3),
                               round(ax * sp.width / emu, 3),
                               round(ay * sp.height / emu, 3)]
        except Exception:
            pass
    if paragraph is not None:
        loc["paragraph"] = paragraph
    if run is not None:
        loc["run"] = run
    if cell is not None:
        loc["cell"] = [int(cell[0]), int(cell[1])]
    if field:
        loc["field"] = True
    if part:
        loc["part"] = part
    return loc or None
