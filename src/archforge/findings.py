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


# Structured data field names per message id, applied positionally to a finding's args
# (0.7). This derives numeric data (effective size, ratio, overflow...) from the args the
# detectors already pass, so agents and non-text reporters key on numbers, not on parsing
# a rendered sentence, with zero changes at the detection sites and zero effect on the
# verdict. A None entry in a tuple skips that arg (e.g. e3's already-rendered autofit note).
_DATA_FIELDS = {
    "e3": ("effective_pt", "hard_min_pt", None),
    "e4": ("tracking",),
    "w1": ("effective_pt", "body_min_pt"),
    "w6": ("cluster_size",),
    "w7": ("contrast_ratio",),
    "w8": ("effective_pt", "small_min_pt"),
    "w9": ("bar_count",),
    "w10": ("page_count",),
    "w11_buzz": ("buzzword_types",),
    "w12": ("house_baseline_in", "off_page_count"),
    "w13": ("effect_count", "page_count"),
    "w14": ("nominal_titles", "total_titles"),
    "w15": ("overlap_pct",),
    "w16": ("overflow_in",),
    "w17": ("inside_pct",),
}


# Auto-generated shape names that carry no identity: PowerPoint / python-pptx assign
# these by shape kind + a counter, so two unrelated textboxes are "TextBox 1"/"TextBox 2"
# and collapse to the same token once the counter is stripped. Google Slides exports use
# "Google Shape;<id>;<page>". For these the bbox is the identity (0.7.1, external review).
_GENERIC_NAME_BASES = frozenset((
    "textbox", "text box", "rectangle", "rounded rectangle", "oval", "ellipse",
    "shape", "freeform", "content placeholder", "text placeholder", "title",
    "subtitle", "picture", "image", "group", "table", "chart", "diagram",
    "straight connector", "elbow connector", "line", "arrow", "placeholder",
    "autoshape", "object", "graphic frame",
))


def _is_generic_shape_name(name: str) -> bool:
    import re
    n = name.strip()
    if n.lower().startswith("google shape"):
        return True
    base = re.sub(r"[\s\d]+$", "", n).strip().lower()
    return base in _GENERIC_NAME_BASES or base == ""


class Finding:
    """A single check result. msg_id+args is canonical and message is a view rendered in the
    current language.

    loc (if present) is a structured target for agent auto-fix: whichever keys are available
    among {"shape_id", "shape_name", "paragraph", "run", "bbox"[x,y,w,h in], "part"}.
    """

    __slots__ = ("page", "code", "msg_id", "args", "detail", "loc", "fp_key",
                 "_msg_override", "_data")

    def __init__(self, page: int, code: str, msg_id: str, args: Tuple = (),
                 detail: str = "", loc: Optional[Dict] = None,
                 fp_key: Optional[str] = None,
                 _msg_override: Optional[str] = None,
                 data: Optional[Dict] = None):
        self.page = page
        self.code = code
        self.msg_id = msg_id
        self.args = tuple(args)
        self.detail = detail
        self.loc = loc
        # Typed payload emitted directly at the detection site (0.8, external review:
        # structured data must come from the detector, not from re-parsing args). The
        # positional _DATA_FIELDS derivation fills anything the site did not state;
        # explicit values win on key collisions.
        self._data = data
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

    def data(self) -> Dict:
        """Structured payload: positional fields derived from args (_DATA_FIELDS, 0.7)
        overlaid with the typed data the detection site emitted directly (0.8). Numbers
        and machine values, never a parsed sentence."""
        out = {}
        fields = _DATA_FIELDS.get(self.msg_id)
        if fields:
            for name, val in zip(fields, self.args):
                if name is None:
                    continue
                out[name] = val
        if self._data:
            out.update(self._data)
        return out

    def to_dict(self, schema: str = "1.0", severity_override: Optional[str] = None) -> Dict:
        d = {"page": self.page, "code": self.code, "message": self.message,
             "detail": self.detail}
        if self.loc:
            d["location"] = self.loc
        if schema != "1.0":
            # v2 adds severity (single findings[] array) and structured data. The
            # effective severity is list membership, not the static registry level:
            # a config severity override can demote E2 to warning (0.8.x), and the
            # reporter must say what actually gated.
            if severity_override:
                d["severity"] = severity_override
            else:
                try:
                    from .rules import severity
                except ImportError:
                    from rules import severity
                d["severity"] = "error" if severity(self.code) == "error" else "warning"
            data = self.data()
            if data:
                d["data"] = data
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

    def _location_bucket(self) -> str:
        """A page-free, coarse structural key for the finding's location (0.7 baseline v3).
        Page is excluded so slide insertion/reorder does not invalidate it. A semantic
        shape name is trusted as identity (regeneration keeps it, position may jitter);
        a generator's auto-name ("TextBox 12", "Rectangle 3", "Google Shape;...") is NOT
        trusted, because stripping its counter collapses every textbox to the same token
        and a defect that moved elsewhere would be silently re-suppressed (0.7.1, external
        review P0). For auto-named or nameless shapes the 0.5in bbox grid is the identity.
        cell/paragraph/field always participate. Empty for locationless deck-level rules."""
        loc = self.loc or {}
        parts = []
        name = loc.get("shape_name")
        if name and not _is_generic_shape_name(str(name)):
            import re
            parts.append("n=" + re.sub(r"[\s\d]+$", "", str(name)).strip().lower())
        elif "bbox" in loc:
            try:
                bx = loc["bbox"]
                parts.append("g=%d,%d" % (round(bx[0] * 2), round(bx[1] * 2)))
            except Exception:
                pass
        if "cell" in loc:
            parts.append("c=%s" % (loc["cell"],))
        if "paragraph" in loc:
            parts.append("p=%s" % loc["paragraph"])
        if loc.get("field"):
            parts.append("fld")
        return "|".join(parts)

    def structural_fp(self) -> str:
        """Baseline v3 fingerprint: the content fingerprint plus the location bucket.
        Stricter than v2 (which pooled same-content findings regardless of place), so a
        defect that disappears and reappears elsewhere is treated as new rather than
        silently re-suppressed (external review). Still page-free."""
        import hashlib
        raw = "%s|%s" % (self.fingerprint(), self._location_bucket())
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
