# -*- coding: utf-8 -*-
"""GitHub Action runner: builds the `archforge scan` argv from env-passed inputs.

Inputs arrive as environment variables set from action inputs, so their values are never
interpolated into a shell script body (0.6.0, external review P0: quotes, semicolons, and
glob characters in inputs used to be interpreted by bash, and `**` was expanded by the
shell before archforge saw it). Multi-value inputs (files, extra-args) are one entry per
line; every entry becomes a single argv element, so paths with spaces survive.
"""
import os
import subprocess
import sys


def _lines(value):
    return [ln.strip() for ln in (value or "").splitlines() if ln.strip()]


def _flag(name, default):
    return (os.environ.get(name) or default).strip().lower() == "true"


def main():
    files = _lines(os.environ.get("AF_FILES"))
    if not files:
        print("archforge action: the 'files' input is empty", file=sys.stderr)
        return 2
    argv = [sys.executable, "-m", "archforge.lint", "scan"] + files
    argv += ["--profile", (os.environ.get("AF_PROFILE") or "core").strip()]
    if not _flag("AF_ALLOW_CONFIG", "false"):
        # Deck-folder configs are untrusted by default in CI (trust boundary)
        argv.append("--no-config")
    if _flag("AF_STRICT", "false"):
        argv.append("--strict")
    else:
        if _flag("AF_FAIL_INCOMPLETE", "true"):
            argv.append("--fail-incomplete")
        if _flag("AF_FAIL_ON_WARNING", "false"):
            argv.append("--fail-on-warning")
    sarif = (os.environ.get("AF_SARIF") or "").strip()
    if sarif:
        argv += ["--sarif", sarif]
    argv += _lines(os.environ.get("AF_EXTRA_ARGS"))
    print("+ " + " ".join(repr(x) if " " in x else x for x in argv[2:]), flush=True)
    return subprocess.call(argv)


if __name__ == "__main__":
    sys.exit(main())
