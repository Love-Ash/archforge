# Recipe: archforge for OfficeCLI output

[OfficeCLI](https://github.com/iOfficeAI/OfficeCLI) is an authoring CLI: agents build
and edit .pptx through DOM paths, with its own HTML engine for previews. archforge is
the gate after it: it does not care who wrote the file, it resolves the font chain and
geometry the way PowerPoint does and returns a deterministic exit code.

The two compose because they check different things. OfficeCLI's preview draws with its
own font logic; whether PowerPoint will resolve a CJK run to a fallback is a property
of PowerPoint's resolver ([../CALIBRATION.md](../CALIBRATION.md)). Concretely, on
OfficeCLI 1.0.135 defaults this deck ships a blocker that no preview shows:

```bash
officecli create deck.pptx
officecli add deck.pptx / --type slide --prop title="Q4 review"
officecli add deck.pptx '/slide[1]' --type shape --prop text="한글 본문" --prop x=1in --prop y=2in
officecli close deck.pptx            # flush before another program reads the file

archforge deck.pptx                  # exit 1: E1, the blank theme's ea slot is empty and
                                     # the run has no a:ea, so Hangul falls back to Malgun
```

Two ways to close it:

```bash
# at authoring time (OfficeCLI supports the East Asian slot directly)
officecli add deck.pptx '/slide[1]' --type shape --prop text="한글 본문" --prop font.ea="Malgun Gothic"

# or after the fact
archforge fix deck.pptx -o deck.pptx --lang en
```

An OfficeCLI fixture family ships in [../../corpus/officecli/](../../corpus/officecli/)
with manifests enforced in CI: the default-theme E1 above, a clean pass with `font.ea`
set, and OfficeCLI-authored reproductions of the other flagship classes (dash
punctuation, sub-5pt text, Hangul tracking damage, frame collision, off-canvas
overflow). None of this can silently regress.

In an agent loop, the pattern is: author with OfficeCLI, then gate before delivery:

```bash
archforge deck.pptx --profile full --fail-incomplete --json   # machine-readable findings
```

Exit 1 blocks the handoff; the JSON carries shape ids and absolute bboxes, so the agent
can patch the exact run through OfficeCLI's own `set` instead of grepping for text.
