# Rendering pages for W7 (contrast over images)

W7 compares text pixels against a rendered PNG of the page, so `--render` needs a folder
of per-page images. Nothing else in archforge needs a renderer, and without `--render` the
report simply marks render-dependent checking as `not_requested` -- it never guesses.

The tool does not render decks itself, deliberately: the target renderer is PowerPoint,
and shipping a different renderer inside the linter would mean judging contrast against
pixels PowerPoint never draws. What you feed `--render` should come from whatever fidelity
you can afford.

## LibreOffice (no PowerPoint required)

```bash
soffice --headless --convert-to pdf deck.pptx --outdir out/
pdftoppm -png -r 96 out/deck.pdf out/p
# out/p-1.png, p-2.png ... -> rename to p01.png, p02.png (zero-padded, page order)
archforge deck.pptx --render out/ --profile full
```

LibreOffice's renderer differs from PowerPoint's (fonts, autofit); treat W7 results from
this path as a screen, not a verdict. `docs/CALIBRATION.md` has the renderer matrix.

## PowerPoint COM (Windows, highest fidelity)

```powershell
$pp = New-Object -ComObject PowerPoint.Application
$deck = $pp.Presentations.Open("C:\path\deck.pptx", $true, $false, $false)
$deck.SaveAs("C:\path\out", 18)   # 18 = ppSaveAsPNG, one PNG per slide
$deck.Close(); $pp.Quit()
```

Rename the exported `슬라이드1.PNG` / `Slide1.PNG` files to `p01.png`-style names.

## Naming contract

`--render DIR` looks for `p01.png`, `p02.png`, ... (zero-padded page numbers). Pages
without a matching PNG are reported as unable-to-check (W18) rather than silently skipped.
