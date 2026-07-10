# -*- coding: utf-8 -*-
"""Per-run Unicode script detection (0.7 decomposition: the parsing layer's most
primitive helpers, pure functions with no dependencies). Re-exported from lint for
backward compatibility."""


def is_hangul(ch):
    o = ord(ch)
    return (0xAC00 <= o <= 0xD7A3      # syllables
            or 0x1100 <= o <= 0x11FF   # jamo
            or 0x3130 <= o <= 0x318F   # compatibility jamo
            or 0xA960 <= o <= 0xA97F   # jamo extended-A (adversarial panel: archaic Hangul
                                       # combinations false negative)
            or 0xD7B0 <= o <= 0xD7FF   # jamo extended-B
            or 0xFFA0 <= o <= 0xFFDC)  # halfwidth Hangul (legacy EDI/OCR output)


def has_hangul(text):
    return any(is_hangul(c) for c in text)


def is_kana(ch):
    return 0x3040 <= ord(ch) <= 0x30FF


def is_hanja(ch):
    o = ord(ch)
    return 0x4E00 <= o <= 0x9FFF or 0x3400 <= o <= 0x4DBF


def is_cjk(ch):
    """Hangul (including extensions), kana, Hanja. Unified script judgment (third review P1:
    fixes the inconsistency where only is_hangul was extended and has_cjk consumers like W8
    were missing the extended Hangul ranges)."""
    return is_hangul(ch) or is_kana(ch) or is_hanja(ch)


def has_cjk(text):
    return any(is_cjk(c) for c in text)


def geometry_unsupported(text):
    """Whether the text contains RTL (Arabic, Hebrew) or complex-shaping scripts (Indic,
    Tibetan, Myanmar, Thai, Lao, Khmer). Because the glyph-width approximation table is a
    Latin/CJK binary, geometry estimates for these scripts are meaningless (0.2.1 script
    layer): rather than guessing, skip and honestly report via W18.
    The range covers 0x0900-0x109F (Indic through Myanmar) plus 0x1780-0x17FF (Khmer)
    (adversarial panel addition)."""
    for c in text:
        o = ord(c)
        if 0x0590 <= o <= 0x08FF or 0x0900 <= o <= 0x109F or 0x1780 <= o <= 0x17FF:
            return True
    return False
