const pptxgen = require("pptxgenjs");

const NAVY = "0F2438";
const NAVY2 = "16324A";
const BG = "F3F6F7";
const INK = "16324A";
const MUTED = "5C7080";
const VIOLET = "8B5CF6";
const BLUE = "2E86FF";
const GREEN = "1FA97E";
const WHITE = "FFFFFF";

const HEAD = "Cambria";
const BODY = "Calibri";

function orbitDecor(slide, cx, cy, color, opacity) {
  // faint concentric orbit rings, molecule motif
  [2.6, 1.9, 1.2].forEach((r, i) => {
    slide.addShape("ellipse", {
      x: cx - r, y: cy - r, w: r * 2, h: r * 2,
      fill: { type: "none" },
      line: { color: color, width: 1.1, transparency: opacity + i * 8 },
    });
  });
  // small nodes on rings
  const nodes = [
    { r: 2.6, ang: 40 }, { r: 1.9, ang: 200 }, { r: 1.2, ang: 300 }, { r: 2.6, ang: 150 },
  ];
  nodes.forEach((n) => {
    const rad = (n.ang * Math.PI) / 180;
    const nx = cx + n.r * Math.cos(rad);
    const ny = cy + n.r * Math.sin(rad);
    slide.addShape("ellipse", {
      x: nx - 0.05, y: ny - 0.05, w: 0.1, h: 0.1,
      fill: { color: color, transparency: 10 }, line: { type: "none" },
    });
  });
}

function footer(slide, label, pageNum, dark) {
  const col = dark ? "8CA0B3" : "8A99A6";
  slide.addText(label.toUpperCase(), {
    x: 0.6, y: 7.08, w: 6, h: 0.3, fontFace: BODY, fontSize: 9,
    color: col, charSpacing: 2, align: "left",
  });
  slide.addText(String(pageNum).padStart(2, "0"), {
    x: 12.13, y: 7.08, w: 0.6, h: 0.3, fontFace: BODY, fontSize: 9,
    color: col, align: "right",
  });
}

function eyebrow(slide, text, color) {
  slide.addShape("rect", { x: 0.6, y: 0.62, w: 0.28, h: 0.1, fill: { color }, line: { type: "none" } });
  slide.addText(text.toUpperCase(), {
    x: 1.0, y: 0.5, w: 8, h: 0.36, fontFace: BODY, fontSize: 13, bold: true,
    color: INK, charSpacing: 2.5,
  });
}

async function main() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
  pres.author = "Oujas Khera";
  pres.title = "Science Portfolio";

  // ---------------- Slide 1: Title ----------------
  let s = pres.addSlide();
  s.background = { color: NAVY };
  orbitDecor(s, 10.6, 5.6, BLUE, 78);
  s.addShape("ellipse", { x: 10.6 - 0.06, y: 5.6 - 0.06, w: 0.12, h: 0.12, fill: { color: VIOLET }, line: { type: "none" } });

  s.addText("SCIENCE", {
    x: 0.75, y: 2.15, w: 10.5, h: 1.15, fontFace: HEAD, bold: true, fontSize: 66,
    color: WHITE, charSpacing: 1,
  });
  s.addText("PORTFOLIO", {
    x: 0.75, y: 3.05, w: 10.5, h: 1.15, fontFace: HEAD, bold: true, fontSize: 66,
    color: BLUE, charSpacing: 1,
  });
  s.addText("A record of experiments, notes & discovery", {
    x: 0.8, y: 4.35, w: 8, h: 0.4, fontFace: BODY, italic: true, fontSize: 15, color: "A9BBCB",
  });

  s.addShape("line", { x: 0.8, y: 5.85, w: 0.7, h: 0, line: { color: BLUE, width: 2 } });
  s.addText("OUJAS KHERA", { x: 0.8, y: 6.0, w: 5, h: 0.34, fontFace: BODY, bold: true, fontSize: 14, color: WHITE, charSpacing: 2 });
  s.addText("GRADE 10E", { x: 0.8, y: 6.32, w: 5, h: 0.3, fontFace: BODY, fontSize: 11, color: "8CA0B3", charSpacing: 2 });

  // legend dots
  const legend = [["CHEMISTRY", VIOLET], ["PHYSICS", BLUE], ["BIOLOGY", GREEN]];
  let lx = 0.8;
  legend.forEach(([label, color]) => {
    s.addShape("ellipse", { x: lx, y: 7.02, w: 0.12, h: 0.12, fill: { color }, line: { type: "none" } });
    s.addText(label, { x: lx + 0.2, y: 6.9, w: 1.6, h: 0.3, fontFace: BODY, fontSize: 10, color: "C4D2DE", charSpacing: 1.5 });
    lx += 1.75;
  });

  // ---------------- Slide 2: Chemistry divider ----------------
  s = pres.addSlide();
  s.background = { color: NAVY };
  orbitDecor(s, 10.7, 3.75, VIOLET, 75);
  s.addShape("ellipse", { x: 0.9, y: 2.15, w: 1.5, h: 1.5, fill: { color: "1E3A57" }, line: { color: VIOLET, width: 1.5 } });
  s.addImage({ path: "icons_science/flask_violet.png", x: 1.32, y: 2.57, w: 0.66, h: 0.66 });
  s.addText("SECTION 01", { x: 0.9, y: 3.85, w: 4, h: 0.35, fontFace: BODY, fontSize: 13, color: VIOLET, bold: true, charSpacing: 3 });
  s.addText("Chemistry", { x: 0.87, y: 4.15, w: 9, h: 1.1, fontFace: HEAD, bold: true, fontSize: 54, color: WHITE });
  s.addText("Reactions, formulae & notebook corrections", { x: 0.9, y: 5.15, w: 8, h: 0.4, fontFace: BODY, italic: true, fontSize: 15, color: "A9BBCB" });
  footer(s, "Science Portfolio", 2, true);

  // ---------------- Slide 3: Chemistry — Activity ----------------
  s = pres.addSlide();
  s.background = { color: BG };
  eyebrow(s, "Chemistry · Activity", VIOLET);
  s.addText("Making Sense of\nChemical Equations", {
    x: 0.6, y: 1.15, w: 5.7, h: 2.0, fontFace: HEAD, bold: true, fontSize: 32, color: INK, lineSpacing: 38,
  });
  s.addText(
    "A hands-on activity balancing chemical equations and mapping ionic reactions — building intuition for how atoms rearrange during a reaction.",
    { x: 0.6, y: 3.15, w: 5.5, h: 1.4, fontFace: BODY, fontSize: 13.5, color: MUTED, lineSpacing: 21 }
  );
  s.addShape("roundRect", { x: 0.6, y: 4.75, w: 0.5, h: 0.5, rectRadius: 0.12, fill: { color: "EDE6FB" }, line: { type: "none" } });
  s.addImage({ path: "icons_science/vial_navy.png", x: 0.72, y: 4.87, w: 0.26, h: 0.26 });
  s.addText("Balanced equation poster · student work", { x: 1.25, y: 4.78, w: 4.5, h: 0.45, fontFace: BODY, fontSize: 11, color: MUTED, italic: true, valign: "middle" });

  s.addShape("roundRect", { x: 6.55, y: 1.0, w: 6.1, h: 5.68, rectRadius: 0.14,
    fill: { color: WHITE }, line: { type: "none" },
    shadow: { type: "outer", color: "1B2A3A", opacity: 0.22, blur: 18, offset: 6, angle: 90 } });
  s.addImage({ path: "crops_science/chem_activity.png", x: 6.75, y: 1.2, w: 5.7, h: 5.28, sizing: { type: "contain", w: 5.7, h: 5.28 } });
  footer(s, "Science Portfolio · Chemistry", 3, false);

  // ---------------- Slide 4: Chemistry — Notebook Corrections ----------------
  s = pres.addSlide();
  s.background = { color: BG };
  eyebrow(s, "Chemistry · Notebook Corrections", VIOLET);
  s.addText("Notebook Corrections", { x: 0.6, y: 1.1, w: 8, h: 0.6, fontFace: HEAD, bold: true, fontSize: 30, color: INK });
  s.addText("Five worked entries revised after teacher feedback, showing corrected steps and reasoning.", {
    x: 0.6, y: 1.62, w: 9, h: 0.45, fontFace: BODY, fontSize: 12.5, color: MUTED,
  });

  const nb = ["chem_nb1", "chem_nb2", "chem_nb3", "chem_nb4", "chem_nb5"];
  const cw = 2.867, ch = 2.15, gap = 0.25;
  const rowY1 = 2.3, rowY2 = 2.3 + ch + gap;
  const startX = (13.333 - (3 * cw + 2 * gap)) / 2;
  for (let i = 0; i < 3; i++) {
    const x = startX + i * (cw + gap);
    s.addShape("roundRect", { x: x - 0.05, y: rowY1 - 0.05, w: cw + 0.1, h: ch + 0.1, rectRadius: 0.1,
      fill: { color: WHITE }, line: { type: "none" }, shadow: { type: "outer", color: "1B2A3A", opacity: 0.16, blur: 10, offset: 3, angle: 90 } });
    s.addImage({ path: `crops_science/${nb[i]}.png`, x, y: rowY1, w: cw, h: ch });
  }
  const startX2 = (13.333 - (2 * cw + gap)) / 2;
  for (let i = 0; i < 2; i++) {
    const x = startX2 + i * (cw + gap);
    s.addShape("roundRect", { x: x - 0.05, y: rowY2 - 0.05, w: cw + 0.1, h: ch + 0.1, rectRadius: 0.1,
      fill: { color: WHITE }, line: { type: "none" }, shadow: { type: "outer", color: "1B2A3A", opacity: 0.16, blur: 10, offset: 3, angle: 90 } });
    s.addImage({ path: `crops_science/${nb[i + 3]}.png`, x, y: rowY2, w: cw, h: ch });
  }
  footer(s, "Science Portfolio · Chemistry", 4, false);

  // ---------------- Slide 5: Physics divider ----------------
  s = pres.addSlide();
  s.background = { color: NAVY };
  orbitDecor(s, 10.7, 3.75, BLUE, 75);
  s.addShape("ellipse", { x: 0.9, y: 2.15, w: 1.5, h: 1.5, fill: { color: "1E3A57" }, line: { color: BLUE, width: 1.5 } });
  s.addImage({ path: "icons_science/atom_blue.png", x: 1.32, y: 2.57, w: 0.66, h: 0.66 });
  s.addText("SECTION 02", { x: 0.9, y: 3.85, w: 4, h: 0.35, fontFace: BODY, fontSize: 13, color: BLUE, bold: true, charSpacing: 3 });
  s.addText("Physics", { x: 0.87, y: 4.15, w: 9, h: 1.1, fontFace: HEAD, bold: true, fontSize: 54, color: WHITE });
  s.addText("Chapters 1–3 · concepts, derivations & problem sets", { x: 0.9, y: 5.15, w: 8, h: 0.4, fontFace: BODY, italic: true, fontSize: 15, color: "A9BBCB" });
  footer(s, "Science Portfolio", 5, true);

  // ---------------- Slide 6: Physics — Chapters 1-3 ----------------
  s = pres.addSlide();
  s.background = { color: BG };
  eyebrow(s, "Physics · Notebook", BLUE);
  s.addText("Chapters 1 – 3", { x: 0.6, y: 1.1, w: 8, h: 0.6, fontFace: HEAD, bold: true, fontSize: 30, color: INK });
  s.addText("Notebook pages spanning the term's first three chapters, corrected and annotated.", {
    x: 0.6, y: 1.62, w: 9, h: 0.4, fontFace: BODY, fontSize: 12.5, color: MUTED,
  });

  const chapters = [["Ch. 1", "phys_ch1"], ["Ch. 2", "phys_ch2"], ["Ch. 3", "phys_ch3"]];
  const pw = 3.75, ph = 2.85, pgap = 0.35;
  const pStartX = (13.333 - (3 * pw + 2 * pgap)) / 2;
  const pY = 2.55;
  chapters.forEach(([label, img], i) => {
    const x = pStartX + i * (pw + pgap);
    s.addShape("roundRect", { x: x - 0.05, y: pY - 0.05, w: pw + 0.1, h: ph + 0.1, rectRadius: 0.1,
      fill: { color: WHITE }, line: { type: "none" }, shadow: { type: "outer", color: "1B2A3A", opacity: 0.18, blur: 12, offset: 4, angle: 90 } });
    s.addImage({ path: `crops_science/${img}.png`, x, y: pY, w: pw, h: ph });
    s.addShape("roundRect", { x: x, y: pY + ph + 0.16, w: 0.95, h: 0.4, rectRadius: 0.08, fill: { color: BLUE }, line: { type: "none" } });
    s.addText(label, { x: x, y: pY + ph + 0.16, w: 0.95, h: 0.4, fontFace: BODY, bold: true, fontSize: 13, color: WHITE, align: "center", valign: "middle" });
  });
  footer(s, "Science Portfolio · Physics", 6, false);

  // ---------------- Slide 7: Biology divider ----------------
  s = pres.addSlide();
  s.background = { color: NAVY };
  orbitDecor(s, 10.7, 3.75, GREEN, 75);
  s.addShape("ellipse", { x: 0.9, y: 2.15, w: 1.5, h: 1.5, fill: { color: "1E3A57" }, line: { color: GREEN, width: 1.5 } });
  s.addImage({ path: "icons_science/leaf_teal.png", x: 1.32, y: 2.57, w: 0.66, h: 0.66 });
  s.addText("SECTION 03", { x: 0.9, y: 3.85, w: 4, h: 0.35, fontFace: BODY, fontSize: 13, color: GREEN, bold: true, charSpacing: 3 });
  s.addText("Biology", { x: 0.87, y: 4.15, w: 9, h: 1.1, fontFace: HEAD, bold: true, fontSize: 54, color: WHITE });

  s.addShape("roundRect", { x: 0.9, y: 5.15, w: 7.6, h: 0.62, rectRadius: 0.1, fill: { color: "17324B" }, line: { color: GREEN, width: 0.75 } });
  s.addText("Notebook currently with teacher for correction — to be added on return.", {
    x: 1.15, y: 5.15, w: 7.2, h: 0.62, fontFace: BODY, italic: true, fontSize: 12.5, color: "CFE3DA", valign: "middle",
  });
  footer(s, "Science Portfolio", 7, true);

  // ---------------- Slide 8: Closing ----------------
  s = pres.addSlide();
  s.background = { color: NAVY };
  orbitDecor(s, 6.67, 3.75, BLUE, 82);
  s.addText("Thank You", { x: 0, y: 3.0, w: 13.333, h: 1.0, fontFace: HEAD, bold: true, fontSize: 44, color: WHITE, align: "center" });
  s.addText("End of Science Portfolio  ·  Oujas Khera, 10E", {
    x: 0, y: 3.9, w: 13.333, h: 0.5, fontFace: BODY, fontSize: 14, color: "A9BBCB", align: "center", charSpacing: 1.5,
  });

  await pres.writeFile({ fileName: "/mnt/user-data/outputs/Science_Portfolio.pptx" });
  console.log("done");
}

main().catch((e) => { console.error(e); process.exit(1); });
