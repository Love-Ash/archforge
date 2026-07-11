# -*- coding: utf-8 -*-
"""Suite-wide fixtures. The Korean-message default keeps assertion text stable
regardless of the host locale; helpers live in helpers.py."""
import os
import sys

import pytest

sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(
    os.path.abspath(__file__))), "src"))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import archforge.messages as jmsg   # noqa: E402


@pytest.fixture(autouse=True)
def _force_korean_messages(monkeypatch):
    """Tests pin the lang to ko to keep existing Korean-message assertions working
    (0.3.0 i18n). English output is verified explicitly in test_lang_*."""
    monkeypatch.setenv("ARCHFORGE_LANG", "ko")
    jmsg.set_lang("ko")
    yield
    jmsg.set_lang(None)
