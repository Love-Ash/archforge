# -*- coding: utf-8 -*-
"""GitHub Action runner: builds the `archforge scan` argv from env-passed inputs.

Inputs arrive as environment variables set from action inputs, so their values are never
interpolated into a shell script body (0.6.0). Multi-value inputs (files, extra-args) are
one entry per line; every entry becomes a single argv element, so paths with spaces
survive and globs are expanded by archforge itself.

0.6.1 (external review): boolean inputs are validated exactly. A typo like
`fail-incomplete: ture` must fail the job, not silently disable a safety default;
"conservatively treat as false" is the opposite of fail-safe for a quality gate. The
runner also emits step outputs and a GITHUB_STEP_SUMMARY table, and supports
changed-only mode (lint only .pptx files changed against a base ref).
"""
import json
import os
import subprocess
import sys
import tempfile


def _lines(value):
    return [ln.strip() for ln in (value or "").splitlines() if ln.strip()]


def _flag(name, env, default):
    raw = (os.environ.get(env) or default).strip().lower()
    if raw not in ("true", "false"):
        print("archforge action: input '%s' must be 'true' or 'false', got %r"
              % (name, raw), file=sys.stderr)
        raise SystemExit(2)
    return raw == "true"


def _changed_pptx(base_ref):
    r = subprocess.run(["git", "diff", "--name-only", "--diff-filter=d",
                        base_ref + "...HEAD"], capture_output=True, text=True)
    if r.returncode != 0:
        print("archforge action: git diff against %r failed (is fetch-depth deep "
              "enough?):\n%s" % (base_ref, r.stderr), file=sys.stderr)
        raise SystemExit(2)
    return [ln for ln in r.stdout.splitlines() if ln.lower().endswith(".pptx")]


def _emit_outputs(report):
    out_path = os.environ.get("GITHUB_OUTPUT")
    if not out_path or not report:
        return
    s = report.get("summary", {})
    with open(out_path, "a", encoding="utf-8") as f:
        f.write("passed=%s\n" % str(bool(s.get("pass"))).lower())
        f.write("error-count=%d\n" % int(s.get("error_count", 0)))
        f.write("warning-count=%d\n" % int(s.get("warn_count", 0)))
        f.write("incomplete=%s\n" % str(bool(s.get("incomplete"))).lower())
        f.write("checked-files=%d\n" % int(s.get("file_count", 0)))
        f.write("failed-files=%d\n" % int(s.get("failed_files", 0)))


def _emit_summary(report):
    path = os.environ.get("GITHUB_STEP_SUMMARY")
    if not path or not report:
        return
    rows = []
    for fdoc in report.get("files", []):
        name = fdoc.get("file", "?")
        if fdoc.get("status") == "error":
            rows.append("| %s | - | - | error: %s |" % (name, fdoc.get("error", "")))
        else:
            s = fdoc.get("summary", {})
            rows.append("| %s | %d | %d | %s |"
                        % (name, s.get("error_count", 0), s.get("warn_count", 0),
                           "incomplete" if s.get("incomplete") else
                           ("fail" if fdoc.get("status") == "fail" else "pass")))
    s = report.get("summary", {})
    with open(path, "a", encoding="utf-8") as f:
        f.write("## archforge\n\n")
        f.write("**%s** - %d file(s), %d failed\n\n"
                % ("PASS" if s.get("pass") else "FAIL",
                   s.get("file_count", 0), s.get("failed_files", 0)))
        f.write("| File | Errors | Warnings | Status |\n|---|---:|---:|---|\n")
        f.write("\n".join(rows) + "\n")


def main():
    changed_only = _flag("changed-only", "AF_CHANGED_ONLY", "false")
    if changed_only:
        base = (os.environ.get("AF_BASE_REF") or "origin/main").strip()
        files = _changed_pptx(base)
        if not files:
            print("archforge action: no .pptx changed against %s; nothing to lint" % base)
            _emit_outputs({"summary": {"pass": True, "error_count": 0, "warn_count": 0,
                                       "incomplete": False, "file_count": 0,
                                       "failed_files": 0}, "files": []})
            return 0
    else:
        files = _lines(os.environ.get("AF_FILES"))
        if not files:
            print("archforge action: the 'files' input is empty", file=sys.stderr)
            return 2

    argv = [sys.executable, "-m", "archforge", "scan"] + files
    argv += ["--profile", (os.environ.get("AF_PROFILE") or "core").strip()]
    if not _flag("allow-config", "AF_ALLOW_CONFIG", "false"):
        # Deck-folder configs are untrusted by default in CI (trust boundary)
        argv.append("--no-config")
    if _flag("strict", "AF_STRICT", "false"):
        argv.append("--strict")
    else:
        if _flag("fail-incomplete", "AF_FAIL_INCOMPLETE", "true"):
            argv.append("--fail-incomplete")
        if _flag("fail-on-warning", "AF_FAIL_ON_WARNING", "false"):
            argv.append("--fail-on-warning")
    if _flag("allow-empty-pattern", "AF_ALLOW_EMPTY_PATTERN", "false"):
        argv.append("--allow-empty-pattern")
    sarif = (os.environ.get("AF_SARIF") or "").strip()
    if sarif:
        argv += ["--sarif", sarif]
    argv += _lines(os.environ.get("AF_EXTRA_ARGS"))

    # Always capture the aggregate JSON for outputs/summary; mirror the human text to
    # the log separately unless the caller asked for JSON themselves via extra-args.
    argv_json = argv + (["--json"] if "--json" not in argv else [])
    print("+ " + " ".join(repr(x) if " " in x else x for x in argv_json[2:]), flush=True)
    with tempfile.TemporaryDirectory() as td:
        out_file = os.path.join(td, "report.json")
        with open(out_file, "w", encoding="utf-8") as f:
            rc = subprocess.call(argv_json, stdout=f)
        report = None
        try:
            with open(out_file, encoding="utf-8") as f:
                report = json.load(f)
        except Exception:
            pass
    if report:
        for fdoc in report.get("files", []):
            if fdoc.get("status") == "error":
                print("ERROR %s: %s" % (fdoc.get("file"), fdoc.get("error")))
            else:
                s = fdoc.get("summary", {})
                print("%-5s %s (errors %d, warnings %d%s)"
                      % (fdoc.get("status", "?").upper(), fdoc.get("file"),
                         s.get("error_count", 0), s.get("warn_count", 0),
                         ", incomplete" if s.get("incomplete") else ""))
        _emit_outputs(report)
        _emit_summary(report)
    return rc


if __name__ == "__main__":
    sys.exit(main())
