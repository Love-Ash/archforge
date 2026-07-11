# Recipe: archforge in a Claude Code deck workflow

The skill pack teaches Claude Code the whole build-lint-fix loop in one command:

```bash
pip install archforge
archforge skill --install        # installs into ./.claude/skills
```

From then on, any session that builds or edits a `.pptx` knows to run

```bash
archforge deck.pptx --profile full --fail-incomplete --json
```

after every build, key on `summary.pass`, and use each finding's `location` payload
(shape id, bbox, paragraph/run) as the auto-fix target instead of searching by text.
For the three mechanically safe rules, let the deterministic fixer do it instead of
the model: `archforge fix deck.pptx -o deck.pptx.fixed && archforge deck.pptx.fixed`.

Tip: add `--html report.html` in the loop and open it when the model claims the deck
is clean; the wireframe shows exactly where anything remaining sits.
