# Recipe: archforge for PptxGenJS output

PptxGenJS writes `fontFace` into the `a:latin` slot and leaves the theme `a:ea` slot
empty, so CJK text on a Latin-only face silently falls back at render time; archforge's
E1 catches exactly this (a PptxGenJS-built fixture ships in [../../corpus/](../../corpus/)).

```bash
node your-deck-generator.js         # produces out/deck.pptx
pip install archforge
archforge out/deck.pptx --profile full --fail-incomplete
```

In a Node project, wire it as a postbuild step:

```json
{ "scripts": { "postbuild": "archforge out/ --profile full --fail-incomplete" } }
```

(`archforge <dir>` scans every .pptx under it; exit 1 blocks the pipeline on defects.)
