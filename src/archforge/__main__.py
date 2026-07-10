# -*- coding: utf-8 -*-
"""`python -m archforge` entry point. Preferred over `python -m archforge.lint`, which
triggers a runpy RuntimeWarning because the package __init__ already imports .lint
(0.6.1, external review)."""
from .lint import main

if __name__ == "__main__":
    main()
