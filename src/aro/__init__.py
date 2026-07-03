# -*- coding: utf-8 -*-
"""aro(아로): 빌드된 .pptx를 배포 전에 기계로 검사하는 한글 특화 품질 린터.

이름은 아로새기다(곱게 새겨 넣다)의 아로에서 왔다. 눈에 안 띄게 조용히 깨지는 한글 타이포를
한 획까지 곱게 짚어 새긴다는 뜻이다.

`aro.lint`는 서브모듈이다(패키지 속성을 함수로 가리지 않는다). 함수형 진입점이
필요하면 `from aro.lint import lint` 또는 `aro.lint_pptx`를 쓴다.
"""
from . import lint as _lint_module  # noqa: F401  (서브모듈 로드 보장)
from .lint import lint as lint_pptx  # noqa: F401
from .lint import frame_autofit, frame_font_scale, main  # noqa: F401

__version__ = "0.1.0"
