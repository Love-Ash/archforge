# -*- coding: utf-8 -*-
"""Regenerates the committed examples/ files. Built from the same source as
`archforge demo` (archforge.demo), so the demo command, repo examples, and
tests share one definition (prevents drift).

Usage: python scripts/make_examples.py
"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "src"))

from archforge import demo   # noqa: E402


def main():
    root = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "examples")
    os.makedirs(root, exist_ok=True)
    jobs = (("broken.pptx", lambda p: demo.build_broken(p, lang="ko")),
            ("fixed.pptx", lambda p: demo.build_fixed(p, lang="ko")),
            ("broken_en.pptx", lambda p: demo.build_broken(p, lang="en")),
            ("fixed_en.pptx", lambda p: demo.build_fixed(p, lang="en")),
            ("style_warnings.pptx", demo.build_warnings))
    for name, builder in jobs:
        path = os.path.normpath(os.path.join(root, name))
        builder(path)
        print("wrote", path)


if __name__ == "__main__":
    main()
