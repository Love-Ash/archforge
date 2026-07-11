# -*- coding: utf-8 -*-
"""This module was split (0.8.x, external audit: a 2,700-line test monolith
mirrors the engine monolith it tests). The suite now lives in:

  helpers.py                 shared deck builders and the CLI runner
  conftest.py                suite-wide fixtures
  test_rules_typography.py   E1-E4 and font resolution
  test_rules_geometry.py     W15-W18 and glyph geometry
  test_rules_structure.py    W6-W14 deck structure and style
  test_baseline_config.py    baseline / config / severity contracts
  test_cli_contracts.py      CLI, scan, and reporter contracts
  test_engine_misc.py        everything else
"""
