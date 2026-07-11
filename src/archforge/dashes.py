# -*- coding: utf-8 -*-
"""E2 dash kernel (0.8, #5 decomposition): the dash-family character class and the
function-based punctuation-vs-range judgment. Pure string functions; no Finding,
no I/O. Re-exported from lint for backward compatibility."""
from typing import List, Optional, Tuple

# Long dash class + fullwidth hyphen-minus + 2x/3x em + box-drawing horizontal line (built
# from code points: no dash character appears literally in this source)
LONG_DASHES = {chr(c) for c in (0x2012, 0x2013, 0x2014, 0x2015, 0x2212, 0xFF0D, 0x2E3A, 0x2E3B, 0x2500)}
_EN_DASH = chr(0x2013)
_MINUS = chr(0x2212)


def _is_digit_ch(c: str) -> bool:
    """Digit judgment covers ASCII and fullwidth digits only: isdigit() also returns True for
    superscript footnote marks (U+00B9) and circled numbers (U+2460), which let a word with a
    footnote mark like "revenue-superscript-1" be mistaken as numeric and slip through the
    exception (measured in the adversarial panel)."""
    return "0" <= c <= "9" or 0xFF10 <= ord(c) <= 0xFF19


def _dash_neighbor(text: str, i: int, step: int) -> Tuple[str, bool]:
    """The neighboring token (up to 12 chars) of the dash at position i, and whether there is
    whitespace. step=-1 for left, +1 for right."""
    j = i + step
    spaced = False
    while 0 <= j < len(text) and text[j].isspace():
        spaced = True
        j += step
    buf = []
    while 0 <= j < len(text) and not text[j].isspace() and text[j] not in LONG_DASHES:
        buf.append(text[j])
        j += step
        if len(buf) >= 12:
            break
    # A left scan (step=-1) accumulates in reverse order, so reverse it back to original text
    # order: this makes the digit-leading check look at the token's actual first character
    # rather than the character adjacent to the dash (fix measured in the adversarial panel)
    tok = "".join(reversed(buf)) if step < 0 else "".join(buf)
    return tok, spaced


def dash_violations(text: str, strict: bool = False,
                    span: Optional[Tuple[int, int]] = None) -> List[str]:
    """Extracts E2 violation characters. The axis of judgment is function, not the character
    itself: it separates range connectors (legitimate typography) from punctuation (an AI
    parenthetical tell) by whether the neighboring tokens are numeric and whether they are
    joined without a space (0.2.1 v2, fixes 4 remaining false-positive shapes from the second
    external re-check).

    en dash (U+2013):
      - both neighboring tokens are numeric (containing a digit: 2020, Q1, 5%, FY24) -> passes
        (a range even with spaces)
      - only one side numeric + the dash is directly adjacent -> passes (a "2020-present"
        style range)
      - only one side numeric + there is a space -> blocked (an AI-style parenthetical like
        "growth - in 2024")
      - word~word -> blocked (cannot be machine-distinguished from a parenthetical, a
        conservative choice. "Seoul~Busan" type cases remain a residual false positive)
    Math minus (U+2212): passes if immediately followed by a digit (negative number/formula).
    em dash and the rest of the dash class: blocked in every context (the signature function).
    strict=True blocks everything with no exceptions.

    If span=(start,end) is given, only characters in that range are reported, but neighboring
    context is read from the whole text. When the caller passes text=whole paragraph and
    span=the run's range, ranges split across run boundaries are not falsely flagged
    (confirmed in the adversarial panel, 2026-07-10)."""
    lo, hi = span if span is not None else (0, len(text))
    bad = []
    for i in range(lo, min(hi, len(text))):
        c = text[i]
        if c not in LONG_DASHES:
            continue
        if not strict:
            # A consecutive dash (an adjacent character before/after is also a dash) has no
            # place in any writing convention: block unconditionally.
            # Closes a hole where, in an exact double sequence like '2020--2024', the two
            # dashes see each other as neighbors, the token becomes empty, and the
            # one-side-numeric exception was slipping through (measured in the adversarial
            # panel).
            prev_ch = text[i - 1] if i > 0 else ""
            next_ch = text[i + 1] if i + 1 < len(text) else ""
            if prev_ch in LONG_DASHES or next_ch in LONG_DASHES:
                bad.append(c)
                continue
            if c == _EN_DASH:
                lt, lsp = _dash_neighbor(text, i, -1)
                rt, rsp = _dash_neighbor(text, i, +1)
                lnum = any(_is_digit_ch(ch) for ch in lt)
                rnum = any(_is_digit_ch(ch) for ch in rt)
                if lnum and rnum:
                    continue
                # The one-side-numeric exception only applies to tokens that start with a
                # digit: closes a bypass where a word+digit mixed token like "conclusion2024"
                # let an adjacent parenthetical pass through (measured in the adversarial
                # panel)
                l_lead = bool(lt) and _is_digit_ch(lt[0])
                r_lead = bool(rt) and _is_digit_ch(rt[0])
                if (l_lead or r_lead) and not lsp and not rsp:
                    continue
            elif c == _MINUS:
                # Allow whitespace: fixes a false positive on financial notation with a spaced
                # sign, like "- 3.2%" (adversarial panel)
                j = i + 1
                while j < len(text) and text[j].isspace():
                    j += 1
                if j < len(text) and _is_digit_ch(text[j]):
                    continue
        bad.append(c)
    return bad
