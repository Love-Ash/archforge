# -*- coding: utf-8 -*-
"""Seed corpus, PowerPoint-native generator: opens python-pptx-built decks in real
PowerPoint via COM and re-saves them, so the corpus includes files written by
PowerPoint's own writer (different XML normalization, relationship layout, and part
ordering than any library). Requires PowerPoint for Windows; run manually, outputs are
committed.

Usage: python corpus/gen_powerpoint_native.py
"""
import json
import os
import shutil
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "python-pptx")
DST = os.path.join(HERE, "powerpoint-native")

RESAVE = ["e1_arial_hangul", "e4_tracked_hangul", "clean_bilingual"]


def main():
    import win32com.client
    os.makedirs(DST, exist_ok=True)
    app = win32com.client.Dispatch("PowerPoint.Application")
    try:
        for name in RESAVE:
            src = os.path.join(SRC, name + ".pptx")
            dst = os.path.join(DST, name + "_native.pptx")
            pres = app.Presentations.Open(src, ReadOnly=True, Untitled=False,
                                          WithWindow=False)
            try:
                pres.SaveCopyAs(dst)
            finally:
                pres.Close()
            with open(os.path.join(SRC, name + ".json"), encoding="utf-8") as f:
                m = json.load(f)
            m["generator"] = "powerpoint-windows (native re-save of python-pptx/%s)" % name
            m["notes"] = (m.get("notes", "") +
                          " | re-saved by PowerPoint's own writer: same content, "
                          "native XML normalization").strip(" |")
            with open(os.path.join(DST, name + "_native.json"), "w",
                      encoding="utf-8") as f:
                json.dump(m, f, ensure_ascii=False, indent=2)
            print("wrote", dst)
    finally:
        app.Quit()


if __name__ == "__main__":
    main()
