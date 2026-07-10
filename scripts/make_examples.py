# -*- coding: utf-8 -*-
"""examples/ 커밋본 재생성. `archforge demo`와 같은 소스(archforge.demo)로 만들어
데모 명령·리포 예제·테스트가 한 정의를 공유한다(드리프트 방지).

사용: python scripts/make_examples.py
"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "src"))

from archforge import demo   # noqa: E402


def main():
    root = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "examples")
    os.makedirs(root, exist_ok=True)
    for name, builder in (("broken.pptx", demo.build_broken),
                          ("fixed.pptx", demo.build_fixed),
                          ("style_warnings.pptx", demo.build_warnings)):
        path = os.path.normpath(os.path.join(root, name))
        builder(path)
        print("wrote", path)


if __name__ == "__main__":
    main()
