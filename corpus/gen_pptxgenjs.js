// Seed corpus, PptxGenJS generator: decks produced by a DIFFERENT generator than the
// one the test suite uses, so cross-generator behavior is covered. Ground truth per
// deck is established by construction plus XML inspection (see corpus/README.md).
//
// Usage: npx -y pptxgenjs@4 --help is not a thing; run via:
//   cd corpus && npm i pptxgenjs@4 --no-save && node gen_pptxgenjs.js
const pptxgen = require("pptxgenjs");
const path = require("path");

const HERE = path.join(__dirname, "pptxgenjs");
const fs = require("fs");
fs.mkdirSync(HERE, { recursive: true });

async function emit(name, build, manifest) {
  const p = new pptxgen();
  p.defineLayout({ name: "WIDE", width: 13.333, height: 7.5 });
  p.layout = "WIDE";
  build(p);
  await p.writeFile({ fileName: path.join(HERE, name + ".pptx") });
  manifest.generator = "pptxgenjs";
  manifest.profile = manifest.profile || "full";
  fs.writeFileSync(path.join(HERE, name + ".json"),
    JSON.stringify(manifest, null, 2) + "\n");
  console.log("wrote", name + ".pptx");
}

(async () => {
  await emit("clean_english", (p) => {
    const s = p.addSlide();
    s.addText("Quarterly results improved across segments", {
      x: 1, y: 1, w: 8, h: 1, fontSize: 20, fontFace: "Calibri",
    });
  }, {
    expected: {},
    notes: "negative fixture: plain English text from PptxGenJS defaults",
  });

  await emit("hangul_fontface_arial", (p) => {
    const s = p.addSlide();
    s.addText("아리알 페이스에 실린 한글", {
      x: 1, y: 1, w: 8, h: 1, fontSize: 20, fontFace: "Arial",
    });
  }, {
    expected: { E1: 1 },
    ground_truth: "by construction + XML inspection: PptxGenJS fontFace writes a:latin " +
      "(and no a:ea), and its default theme carries an empty ea slot, so the measured " +
      "model resolves Arial for the Hangul run",
    notes: "the cross-generator twin of python-pptx e1_arial_hangul",
  });

  await emit("em_dash_prose", (p) => {
    const s = p.addSlide();
    s.addText("growth was structural—not cyclical", {
      x: 1, y: 1, w: 8, h: 1, fontSize: 16, fontFace: "Calibri",
    });
  }, {
    expected: { E2: 1 },
  });
})();
