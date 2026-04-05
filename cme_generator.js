/*
  CME PRESENTATION GENERATOR
  Collaborate Middle East — Internal Tool

  HOW TO USE:
  1. Edit the BRIEF object below with your presentation content
  2. Run: node cg_generator.js
  3. Open cme_output.pptx from the output folder

  Available slide types: title, divider, sub_divider, contents, headline_body,
  key_stats, blue_cards, four_column, timeline, single_image, two_images,
  venn, quote, team_cards, thank_you, challenge, flow_diagram, end

  Do not edit anything below the BRIEF object unless you are extending the script.
*/

const pptxgen = require("pptxgenjs");
const fs      = require("fs");
const path    = require("path");

// --- HARDCODED PATHS --- //
const ASSETS_DIR = "/home/claude/cme_assets";
const OUTPUT_PATH = "/home/claude/cme_output.pptx";

// --- BRAND CONSTANTS --- //
const FONT = "General Sans";  // All text throughout the document
const C = {
  black:     "000000",
  dark:      "3C3830",
  white:     "FFFFFF",
  offWhite:  "EDE8DC",
  cream:     "C5BDA4",
  sand:      "8B8470",
  sandDark:  "5C5648",
  gold:      "C4943A",
  goldLight: "D4B870",
  lightGold: "EAD89C",
  linkGreen: "3D6B4F",
  grey:      "EAD89C",
  midGrey:   "8B8470",
  blueCard:  "D4B870",
  lightBlue: "C5BDA4",
};
const W = 13.33;
const H = 7.5;
const ML = 0.38; // master left margin — every text/line starts here

// --- IMAGE LOADER --- //
function loadImage(filename) {
  const ext  = path.extname(filename).slice(1).toLowerCase();
  const mime = (ext === "jpg" || ext === "jpeg") ? "image/jpeg" : "image/png";
  const data = fs.readFileSync(path.join(ASSETS_DIR, filename));
  return `${mime};base64,${data.toString("base64")}`;
}

// --- LOGO RENDERER --- //
// Logos are rendered at runtime from inline SVG source via cairosvg.
// This bypasses the project PNG files entirely — the output is always a perfect
// RGBA PNG with clean transparent background, regardless of what PNG files exist
// in the assets folder. The SVG paths are the single source of truth.
function renderLogoFromSVG(svgContent, scale, fallbackFile) {
  const { execSync } = require("child_process");
  const escaped = svgContent.replace(/\\/g, "\\\\").replace(/'/g, "\\'").replace(/\n/g, "\\n");
  const script = `
import cairosvg, sys
svg = b'''${escaped}'''
png = cairosvg.svg2png(bytestring=svg, scale=${scale}, background_color=None)
sys.stdout.buffer.write(png)
`;
  try {
    const pngBuffer = execSync(`python3 -c "${script.replace(/"/g, '\\"')}"`, { maxBuffer: 20 * 1024 * 1024 });
    return `image/png;base64,${pngBuffer.toString("base64")}`;
  } catch(e) {
    // cairosvg not available — fall back to pre-rendered PNG from assets folder
    if (fallbackFile) {
      const data = fs.readFileSync(path.join(ASSETS_DIR, fallbackFile));
      console.log(`Logo fallback: using ${fallbackFile} (cairosvg unavailable)`);
      return `image/png;base64,${data.toString("base64")}`;
    }
    throw new Error(`renderLogoFromSVG failed and no fallback specified: ${e.message}`);
  }
}

// CME logos loaded directly from PNG files (SVG source too large to embed)
console.log("Loading CME logos from PNG assets...");
const logoBlack  = `image/png;base64,${fs.readFileSync(path.join(ASSETS_DIR, 'logo_black.png')).toString('base64')}`;
const logoWhite  = `image/png;base64,${fs.readFileSync(path.join(ASSETS_DIR, 'logo_white.png')).toString('base64')}`;
const xIconBlack = `image/png;base64,${fs.readFileSync(path.join(ASSETS_DIR, 'x_icon_black.png')).toString('base64')}`;
const xIconWhite = `image/png;base64,${fs.readFileSync(path.join(ASSETS_DIR, 'x_icon_white.png')).toString('base64')}`;
console.log("Logos ready.");

// --- BRAND IMAGES --- //
const imgPlaceholder = loadImage("img_placeholder.png"); // neutral grey — use for all image slots
const phoneHand      = loadImage("page11_img2.jpeg");    // hand holding iPhone, white screen
const titleBg        = loadImage("title_bg.jpeg");      // Title slide — geometric gradient
const gradientFull   = loadImage("gradient_full.jpeg"); // Full-slide soft gradient (divider)
const gradientHalf   = loadImage("gradient_half.jpeg"); // Bottom-half gradient (sub-divider)
const gradientStrip  = loadImage("gradient_strip.jpeg");// Thin bottom strip (content slides)
const gradientPanel  = loadImage("gradient_panel.jpeg");// Right panel gradient (split slides)
const gradientWide   = loadImage("gradient_wide.jpeg"); // Wide half gradient (team, four-col)
const endBg          = loadImage("end_bg.jpeg");        // End slide gradient

// --- HELPER FUNCTIONS --- //
// These must be used on every slide — never hardcode positions manually

function addLogoBlack(slide) {
  // Black logo — top-left, sized to match template proportions
  slide.addImage({ data: logoBlack, x: 0.29, y: 0.18, w: 1.38, h: 0.34 });
}

function addLogoWhite(slide) {
  // White logo — top-left
  slide.addImage({ data: logoWhite, x: 0.29, y: 0.18, w: 1.38, h: 0.34 });
}

function addXBlack(slide) {
  // Black X icon top-right corner
  slide.addImage({ data: xIconBlack, x: 12.9, y: 0.2, w: 0.24, h: 0.24 });
}

function addXWhite(slide) {
  // White X icon top-right corner
  slide.addImage({ data: xIconWhite, x: 12.9, y: 0.2, w: 0.24, h: 0.24 });
}

function addSectionLabel(slide, text, opts = {}) {
  // Small section label top-left (dark text — white bg slides only)
  const color = opts.white ? C.white : C.dark;
  slide.addText(text, {
    x: 0.29, y: 0.15, w: 3.0, h: 0.28,
    fontSize: 10, fontFace: FONT, color, bold: false
  });
}

function addHeadline(slide, parts, opts = {}) {
  // parts: array of { text, bold } objects
  // opts: x, y, w, h overrides
  const x = opts.x !== undefined ? opts.x : 0.29;
  const y = opts.y !== undefined ? opts.y : 0.75;
  const w = opts.w !== undefined ? opts.w : 4.5;
  const h = opts.h !== undefined ? opts.h : 1.8;
  const size = opts.size || 32;
  const color = opts.color || C.dark;

  const runs = parts.map(p => ({
    text: p.text,
    options: { bold: !!p.bold, fontSize: size, fontFace: FONT, color }
  }));
  slide.addText(runs, { x, y, w, h, valign: "top", margin: 0 });
}

// ============================================================
// SLIDE TYPE FUNCTIONS
// ============================================================

// --- FULL-BLEED IMAGE SLIDE --- //
// Full-bleed photo, X icon top-right, large bold white headline bottom-left
// When no image is supplied, a grey placeholder fills the slide and image
// guidance (ChatGPT prompt + Google search suggestion + ratio) is printed to console.
function addFullBleedSlide(pres, data) {
  const slide = pres.addSlide();

  if (data.image) {
    slide.addImage({ data: data.image, x: 0, y: 0, w: W, h: H,
      sizing: { type: "cover", w: W, h: H } });
  } else {
    // Placeholder — dark grey so white text is readable
    slide.addShape(pres.ShapeType.rect, {
      x: 0, y: 0, w: W, h: H,
      fill: { color: "2A2A2A" }, line: { color: "2A2A2A" }
    });
    // Print image guidance to console for the user
    const ratio   = data.imageRatio || "16:9";
    const prompt  = data.imagePrompt || "A dramatic, high-quality brand experience or live event photograph. Dark, moody atmosphere with vibrant accent lighting. Premium, editorial feel.";
    const google  = data.googleSearch || "site:unsplash.com brand activation live event dramatic lighting";
    console.log(`\n[IMAGE PLACEHOLDER] Slide: "${data.headline || "Full bleed"}"`);
    console.log(`  Ratio:         ${ratio}`);
    console.log(`  ChatGPT prompt: "${prompt}"`);
    console.log(`  Google search:  ${google}\n`);

    // Show guidance text on the placeholder itself
    slide.addText([
      { text: "IMAGE PLACEHOLDER\n", options: { bold: true, fontSize: 14, color: C.white } },
      { text: `Ratio: ${ratio}\n\n`, options: { fontSize: 11, color: C.white } },
      { text: `ChatGPT prompt:\n`, options: { bold: true, fontSize: 10, color: C.white } },
      { text: `${prompt}\n\n`, options: { fontSize: 10, color: C.white, italic: true } },
      { text: `Google search:\n`, options: { bold: true, fontSize: 10, color: C.white } },
      { text: google, options: { fontSize: 10, color: C.white, italic: true } }
    ], {
      x: 4.0, y: 1.5, w: 5.0, h: 4.5,
      fontFace: FONT, valign: "top", margin: 0
    });
  }

  // Section label top-left
  addSectionLabel(slide, data.section || "", { white: true });
  // X icon top-right (white — dark background)
  addXWhite(slide);

  // Large bold white headline bottom-left
  if (data.headline) {
    const words = data.headline.split(" ");
    // Bold the key word (last word by default, or data.boldWord if specified)
    const boldWord = data.boldWord || words[words.length - 1];
    const runs = [];
    data.headline.split(" ").forEach((word, i) => {
      if (i > 0) runs.push({ text: " ", options: { fontSize: 64, fontFace: FONT, color: C.white } });
      runs.push({
        text: word,
        options: { bold: word === boldWord, fontSize: 64, fontFace: FONT, color: C.white }
      });
    });
    slide.addText(runs, {
      x: 0.29, y: 3.8, w: 11.0, h: 3.4,
      valign: "bottom", margin: 0
    });
  }
}

// --- TITLE SLIDE --- //
// HEADLINE RULE: exactly one line bold + one line regular — never more.
// The AI writing the BRIEF must choose a title that fits this constraint.
// Text box is 3/4 page width (10.0") to allow longer titles without wrapping.
// Font size 72pt — do not reduce; if title wraps, shorten the copy instead.
function addTitleSlide(pres, data) {
  const slide = pres.addSlide();
  slide.addShape(pres.ShapeType.rect, {
    x: 0, y: 0, w: W, h: H, fill: { color: "C7D2FE" }, line: { color: "C7D2FE" }
  });
  slide.addImage({ data: titleBg, x: 0, y: 0, w: W, h: H });

  // Top bar: label centre, date top-right
  if (data.label) {
    slide.addText(data.label, {
      x: 4.5, y: 0.26, w: 4.0, h: 0.30,
      fontSize: 11, fontFace: FONT, color: C.dark, align: "center"
    });
  }
  if (data.date) {
    slide.addText(data.date, {
      x: 10.0, y: 0.26, w: 3.04, h: 0.30,
      fontSize: 11, fontFace: FONT, color: C.dark, align: "right"
    });
  }

  addLogoBlack(slide);

  // Headline — 3/4 page width (10.0"), line 1 bold, line 2 regular
  // One line each — AI must ensure no wrapping at 72pt
  const lines = (data.headline || "Title\nSubtitle").split("\n");
  const runs = [];
  lines.forEach((line, i) => {
    if (i > 0) runs.push({ text: "\n", options: { fontSize: 72, fontFace: FONT, color: C.dark } });
    runs.push({
      text: line,
      options: { bold: i === 0, fontSize: 72, fontFace: FONT, color: C.dark }
    });
  });
  slide.addText(runs, {
    x: 0.29, y: 3.2, w: 10.0, h: 4.0,
    valign: "bottom", margin: 0
  });
}

// --- END / LOGO SLIDE --- //
function addEndSlide(pres) {
  const slide = pres.addSlide();
  slide.addShape(pres.ShapeType.rect, {
    x: 0, y: 0, w: W, h: H, fill: { color: "C7D2FE" }, line: { color: "C7D2FE" }
  });
  slide.addImage({ data: endBg, x: 0, y: 0, w: W, h: H });
  // Logo centred: image is 500px wide with X icon ~80px + gap + wordmark
  // At w=3.8" the rendered wordmark starts at ~x=4.77 to centre it
  slide.addImage({ data: logoBlack, x: 4.77, y: 3.28, w: 3.8, h: 0.95 });
}

// --- HEADLINE + BODY SLIDE --- //
function addHeadlineBodySlide(pres, data) {
  const slide = pres.addSlide();

  // White background
  // Thin gradient strip at bottom
  slide.addImage({ data: gradientStrip, x: 0, y: 6.92, w: W, h: 0.806 });

  addSectionLabel(slide, data.section || "");
  addXBlack(slide);

  // Headline: "The [bold] rest" — tighter to top, matching template
  addHeadline(slide, [
    { text: "The ", bold: false },
    { text: data.boldWord || "", bold: true },
    { text: " " + (data.rest || ""), bold: false }
  ], { x: 0.29, y: 0.52, w: 6.0, h: 1.7, size: 32 });

  // Body text — immediately below headline
  if (data.body) {
    slide.addText(data.body, {
      x: 0.29, y: 2.38, w: 12.5, h: 4.3,
      fontSize: 14, fontFace: FONT, color: C.dark,
      valign: "top", margin: 0
    });
  }
}

// Helper: parse inline [bold]...[/bold] markers into PptxGenJS runs
// IMPORTANT: no spurious line breaks are inserted at the bold/regular boundary.
// Each [bold]...[/bold] block is a single run; surrounding text is split on \n
// so that explicit paragraph breaks in copy are preserved but no extra breaks added.
function parseInlineBold(text, fontSize, color) {
  const runs = [];
  const parts = String(text).split(/(\[bold\][\s\S]*?\[\/bold\])/g);
  parts.forEach(part => {
    const m = part.match(/\[bold\]([\s\S]*?)\[\/bold\]/);
    if (m) {
      // Bold run — emit as single run, no surrounding newlines added
      runs.push({ text: m[1], options: { bold: true, fontSize, fontFace: FONT, color } });
    } else if (part) {
      // Regular text — split on explicit \n only
      const lines = part.split("\n");
      lines.forEach((line, li) => {
        if (li > 0) runs.push({ text: "\n", options: { bold: false, fontSize, fontFace: FONT, color } });
        if (line) runs.push({ text: line, options: { bold: false, fontSize, fontFace: FONT, color } });
      });
    }
  });
  return runs;
}

// --- THREE-COLUMN BODY SLIDE --- //
// White bg, headline top-left, three columns with X chevron accent + body copy
// Thin gradient strip at bottom. Matches template page 5/6.
function addThreeColumnSlide(pres, data) {
  const slide = pres.addSlide();

  slide.addImage({ data: gradientStrip, x: 0, y: 6.92, w: W, h: 0.806 });

  addSectionLabel(slide, data.section || "");
  addXBlack(slide);

  // Line 1 bold, line 2 regular — tight box, aligned to master left margin
  addHeadline(slide, [
    { text: data.headline || "", bold: true },
    { text: (data.headlineBold ? "\n" + data.headlineBold : ""), bold: false }
  ], { x: ML, y: 0.48, w: 7.0, h: 1.55, size: 28 });

  const cols = data.columns || [];
  // Each column width = line width — tight to the text, not full-slide thirds
  const lineW   = 4.0;   // width of the accent line and text column
  const colGap  = 0.35;  // gap between columns
  const col0X   = ML;

  cols.slice(0, 3).forEach((col, i) => {
    const x  = col0X + i * (lineW + colGap);
    const cw = lineW;

    // Heavier accent line — same width as text column, no chevron
    slide.addShape(pres.ShapeType.line, {
      x, y: 2.42, w: cw, h: 0,
      line: { color: C.dark, width: 1.5 }
    });

    // Body paragraphs — start just below line, width = line width
    const paras = Array.isArray(col.body) ? col.body : [col.body];
    let yOffset = 2.56;
    paras.forEach(para => {
      const runs = parseInlineBold(para, 11.5, C.dark);
      slide.addText(runs, {
        x, y: yOffset, w: cw, h: 1.6,
        valign: "top", margin: 0
      });
      yOffset += 1.65;
    });

    // Optional attribution
    if (col.attribution) {
      const attrRuns = parseInlineBold(col.attribution, 11, C.dark);
      slide.addText(attrRuns, {
        x, y: 6.05, w: cw, h: 0.6,
        valign: "top", margin: 0
      });
    }
  });
}

// --- CHALLENGE / BRIEF SLIDE --- //
// Full gradient bg, two-column layout: left = headline, right = Challenge/Idea/Experience labels
// Matches template page 36
function addChallengeSlide(pres, data) {
  const slide = pres.addSlide();
  const splitX = 4.44; // left white ~1/3, right gradient ~2/3

  // White left panel
  slide.addShape(pres.ShapeType.rect, {
    x: 0, y: 0, w: splitX, h: H, fill: { color: C.white }, line: { color: C.white }
  });

  // Gradient right panel
  slide.addImage({ data: gradientPanel, x: splitX, y: 0, w: W - splitX, h: H });

  addSectionLabel(slide, data.section || "");
  addXBlack(slide);

  // Headline — line 1 bold, line 2 regular — bottom-left of left panel
  addHeadline(slide, [
    { text: data.headline || "", bold: true },
    { text: (data.headlineBold ? "\n" + data.headlineBold : ""), bold: false }
  ], { x: ML, y: 3.2, w: splitX - ML - 0.2, h: 3.8, size: 32, color: C.dark });

  // Right side: labelled sections — on gradient
  const sections = data.sections || [];
  let yPos = 1.1;
  sections.forEach(sec => {
    slide.addText(sec.label + ":", {
      x: splitX + 0.4, y: yPos, w: 8.2, h: 0.3,
      fontSize: 12, fontFace: FONT, bold: true, color: C.dark, valign: "top", margin: 0
    });
    slide.addText(sec.body, {
      x: splitX + 0.4, y: yPos + 0.33, w: 8.2, h: 1.05,
      fontSize: 11.5, fontFace: FONT, color: C.dark, valign: "top", margin: 0
    });
    yPos += 1.52;
  });
}

// --- FLOW DIAGRAM SLIDE --- //
// JSX-inspired 2026 styling: numbered steps connected by arrows, micro-labels,
// gradient-tinted cards with subtle borders — matches the Sales Intelligence engine's
// card design language (frosted, bordered, layered hierarchy).
// Layout: up to 6 nodes. 1–3 = single row; 4–6 = two rows of 3.
// FLOW DIAGRAM HEADLINE RULE: maximum two lines (one bold + one regular).
// AI must keep titles short enough to fit at 28pt in a 6.6" wide box.
// Text box extends to half-page width to accommodate longer titles without wrapping.
function addFlowDiagramSlide(pres, data) {
  const slide = pres.addSlide();

  // Bottom gradient strip
  slide.addImage({ data: gradientStrip, x: 0, y: 6.92, w: W, h: 0.806 });

  addSectionLabel(slide, data.section || "");
  addXBlack(slide);

  // Headline box extended to W/2 (6.665") — prevents three-line wrap on longer titles
  addHeadline(slide, [
    { text: data.boldWord || "", bold: true },
    { text: "\n" + (data.rest || ""), bold: false }
  ], { x: ML, y: 0.48, w: W / 2, h: 1.3, size: 28 });

  const nodes = data.nodes || [];
  const count = nodes.length;
  if (count === 0) return;

  // Determine layout: single row or 2 rows of 3
  const rows     = count <= 3 ? 1 : 2;
  const perRow   = Math.ceil(count / rows);
  const cardW    = (W - ML - 0.2) / perRow - 0.18;
  const cardH    = rows === 1 ? 3.2 : 2.0;
  const startY   = rows === 1 ? 2.2 : 1.65;
  const rowGap   = 0.35;

  // Subtle tint colours cycling per card — from JSX palette
  const tints = [
    { fill: "EEF1FF", border: "B8C4E8" },  // blue-lavender
    { fill: "F0F4FF", border: "C4CEFF" },  // periwinkle
    { fill: "F5F0FF", border: "CBBFEE" },  // violet
  ];

  nodes.forEach((node, i) => {
    const row  = Math.floor(i / perRow);
    const col  = i % perRow;
    const x    = ML + col * (cardW + 0.18);
    const y    = startY + row * (cardH + rowGap);

    const tint = tints[col % tints.length];

    // Card background — frosted tint with border
    slide.addShape(pres.ShapeType.rect, {
      x, y, w: cardW, h: cardH,
      fill: { color: tint.fill },
      line: { color: tint.border, width: 1.0 }
    });

    // Step number pill — top-left corner of card
    const numSize = 0.32;
    slide.addShape(pres.ShapeType.ellipse, {
      x: x + 0.15, y: y + 0.15, w: numSize, h: numSize,
      fill: { color: C.dark }, line: { color: C.dark, width: 0 }
    });
    slide.addText(String(i + 1), {
      x: x + 0.15, y: y + 0.15, w: numSize, h: numSize,
      fontSize: 9, fontFace: FONT, bold: true, color: C.white,
      align: "center", valign: "middle", margin: 0
    });

    // Node title (if provided) — micro-label style, uppercase, tight tracking
    if (node.title) {
      slide.addText(node.title.toUpperCase(), {
        x: x + 0.15, y: y + 0.56, w: cardW - 0.3, h: 0.24,
        fontSize: 7.5, fontFace: FONT, bold: true, color: "7B90D4",
        valign: "top", margin: 0
      });
    }

    // Main text
    const textY = node.title ? y + 0.84 : y + 0.58;
    const textH = cardH - (node.title ? 0.84 : 0.58) - 0.2;
    slide.addText(node.text || "", {
      x: x + 0.15, y: textY, w: cardW - 0.3, h: textH,
      fontSize: rows === 1 ? 12 : 11, fontFace: FONT, color: C.dark,
      valign: "top", margin: 0
    });

    // Connector arrow to next card (same row only, not after last in row)
    const isLastInRow = (col === perRow - 1) || (i === count - 1);
    if (!isLastInRow) {
      const arrowX = x + cardW + 0.04;
      const arrowY = y + cardH / 2 - 0.12;
      slide.addText("›", {
        x: arrowX, y: arrowY, w: 0.14, h: 0.24,
        fontSize: 22, fontFace: FONT, bold: true, color: "B8C4E8",
        align: "center", margin: 0
      });
    }
  });
}

// --- TEAM CARDS SLIDE --- //
// Shows 3–5 team members per slide. If more than 5 members are supplied,
// automatically generates two consecutive slides splitting members evenly.
// Fields: name (12pt bold), title (10pt bold), bio (10pt regular, 2-3 sentences max).
// Phone and email removed. Grey gradient photo placeholder with magenta headshot instruction.
// DYNAMIC: cards resize proportionally based on member count (3, 4 or 5 per slide).
function addTeamCardsSlide(pres, data) {
  const allMembers = data.members || [];
  const MAX_PER_SLIDE = 5;

  // Split into pages if needed
  const pages = [];
  for (let i = 0; i < allMembers.length; i += MAX_PER_SLIDE) {
    pages.push(allMembers.slice(i, i + MAX_PER_SLIDE));
  }
  if (pages.length === 0) pages.push([]);

  pages.forEach((members, pageIdx) => {
    const slide  = pres.addSlide();
    const splitY = 3.75;
    const count  = Math.max(members.length, 1);

    // Gradient bottom half (background behind cards)
    slide.addImage({ data: gradientWide, x: 0, y: splitY, w: W, h: H - splitY });

    addSectionLabel(slide, data.section || "");
    addXBlack(slide);

    // Headline — show page indicator if multiple slides
    const headText = pages.length > 1
      ? `Your ${data.boldWord || "team"} (${pageIdx + 1}/${pages.length})`
      : (data.boldWord || "team");
    addHeadline(slide, [
      { text: "Your ", bold: false },
      { text: headText, bold: true }
    ], { x: 0.29, y: 0.48, w: 6.0, h: 0.85, size: 32 });

    const leftPad  = ML;
    const rightPad = ML + 0.1;
    const cardGap  = 0.12;
    const totalW   = W - leftPad - rightPad - (cardGap * (count - 1));
    const cardW    = totalW / count;
    const imgTopY  = 1.55;
    const imgH     = splitY - imgTopY;

    members.forEach((m, i) => {
      const x  = leftPad + i * (cardW + cardGap);
      const cw = cardW;

      // Photo placeholder — grey gradient, right-click → Change Picture in PowerPoint
      if (m.image) {
        slide.addImage({ data: m.image, x, y: imgTopY, w: cw, h: imgH,
          sizing: { type: "cover", w: cw, h: imgH } });
      } else {
        slide.addImage({ data: gradientPanel, x, y: imgTopY, w: cw, h: imgH,
          sizing: { type: "cover", w: cw, h: imgH } });
        // Magenta headshot instruction — unmissable, must be deleted after use
        slide.addText("Insert headshot here", {
          x: x + 0.1, y: imgTopY + imgH / 2 - 0.12, w: cw - 0.2, h: 0.26,
          fontSize: 9, fontFace: FONT, bold: true, color: "EC4899",
          align: "center", valign: "middle", margin: 0
        });
      }

      // White card background bottom half
      slide.addShape(pres.ShapeType.rect, {
        x, y: splitY, w: cw, h: H - splitY,
        fill: { color: C.white }, line: { color: "D0D8EE", width: 0.5 }
      });

      // Text starts below the splitY line (where white card begins)
      const cardTop = splitY + 0.16;
      const cardBot = H - 0.15;
      const cardTxtH = cardBot - cardTop;

      // Name — 12pt bold
      slide.addText(m.name || "Name", {
        x: x + 0.1, y: cardTop, w: cw - 0.2, h: 0.34,
        fontSize: 12, fontFace: FONT, bold: true,
        color: C.dark, valign: "top", margin: 0
      });

      // Job title — 10pt bold
      slide.addText(m.title || "", {
        x: x + 0.1, y: cardTop + 0.36, w: cw - 0.2, h: 0.28,
        fontSize: 10, fontFace: FONT, bold: true,
        color: C.dark, valign: "top", margin: 0
      });

      // Bio — 10pt regular, fills remaining card space
      slide.addText(m.bio || "", {
        x: x + 0.1, y: cardTop + 0.68, w: cw - 0.2, h: cardTxtH - 0.68,
        fontSize: 10, fontFace: FONT, color: C.dark,
        valign: "top", margin: 0
      });
    });
  });
}


// --- BLUE CARDS SLIDE --- //
// White top with headline, 4 blue-card columns below — each card has number, title, body, bullets
// Matches template page 16
function addBlueCardsSlide(pres, data) {
  const slide = pres.addSlide();

  addSectionLabel(slide, data.section || "");
  addXBlack(slide);

  addHeadline(slide, [
    { text: "The " + (data.boldWord || ""), bold: true },
    { text: "\n" + (data.rest || ""), bold: false }
  ], { x: ML, y: 0.48, w: 6.0, h: 1.55, size: 32 });

  const cards = data.cards || [];
  const count  = cards.length;
  const cardW  = (W - ML - 0.1) / Math.max(count, 1);

  // Single gradient band across lower third of slide
  const gradY = 4.7;
  slide.addImage({ data: gradientWide, x: 0, y: gradY, w: W, h: H - gradY });

  const titleY  = 2.6;
  const bodyY   = titleY + 0.5;
  const addlY   = gradY + 0.15;

  cards.forEach((card, i) => {
    const x  = ML + i * cardW;
    const cw = cardW - 0.15;

    // Number + title bold — all start at ML
    slide.addText(`${i + 1}.  ${card.title}`, {
      x, y: titleY, w: cw, h: 0.45,
      fontSize: 13, fontFace: FONT, bold: true,
      color: C.dark, valign: "top", margin: 0
    });

    // Body copy
    slide.addText(card.body || "", {
      x, y: bodyY, w: cw, h: 2.05,
      fontSize: 11, fontFace: FONT, color: C.dark,
      valign: "top", margin: 0
    });

    // Additional bullets sit on gradient band
    if (card.additional && card.additional.length) {
      slide.addText("Additional:", {
        x, y: addlY, w: cw, h: 0.26,
        fontSize: 10, fontFace: FONT, bold: true,
        color: C.dark, valign: "top", margin: 0
      });
      const bulletText = card.additional.map(b => "• " + b).join("\n");
      slide.addText(bulletText, {
        x, y: addlY + 0.3, w: cw, h: H - addlY - 0.35,
        fontSize: 10, fontFace: FONT, color: C.dark, italic: true,
        valign: "top", margin: 0
      });
    }
  });
}

// --- QUOTE SLIDE --- //
// Full gradient background, large centred quote text, attribution below
function addQuoteSlide(pres, data) {
  const slide = pres.addSlide();
  slide.addShape(pres.ShapeType.rect, { x: 0, y: 0, w: W, h: H, fill: { color: "BFC8F5" }, line: { color: "BFC8F5" } });
  slide.addImage({ data: gradientFull, x: 0, y: -0.123, w: W, h: 7.746 });
  addXBlack(slide);
  addSectionLabel(slide, data.section || "", { white: false });

  // Opening mark — separate text box, large, top-left
  slide.addText("\u201C", {
    x: ML, y: 1.8, w: 1.2, h: 1.0,
    fontSize: 96, fontFace: FONT, color: C.dark,
    valign: "top", margin: 0
  });

  // Quote body — italic, generous box
  slide.addText(data.quote || "", {
    x: ML, y: 2.65, w: 12.5, h: 2.5,
    fontSize: 24, fontFace: FONT, color: C.dark,
    italic: true, valign: "top", margin: 0
  });

  // Closing mark — immediately after quote text, right-aligned
  slide.addText("\u201D", {
    x: ML, y: 5.1, w: 12.5, h: 0.9,
    fontSize: 96, fontFace: FONT, color: C.dark,
    align: "right", valign: "top", margin: 0
  });

  if (data.attribution) {
    slide.addText(data.attribution, {
      x: ML, y: 6.1, w: 12.5, h: 0.5,
      fontSize: 12, fontFace: FONT, color: C.dark,
      valign: "top", margin: 0
    });
  }
}

// --- TWO IMAGES SLIDE --- //
// Same left panel as single image, but right side splits into two stacked images
function addTwoImagesSlide(pres, data) {
  const slide = pres.addSlide();
  const splitX = 3.55;
  const imgW   = W - splitX;
  const halfH  = H / 2;

  // White left panel with gradient lower portion
  slide.addShape(pres.ShapeType.rect, {
    x: 0, y: 0, w: splitX, h: H, fill: { color: C.white }, line: { color: C.white }
  });
  // Left panel fully white — no gradient patch

  // Thin white gutter between the two images (0.06" = ~2px at 96dpi)
  const gutterH = 0.06;
  const imgHalf = (H - gutterH) / 2;

  // Top image placeholder
  addImagePlaceholder(slide, splitX, 0, imgW, imgHalf, {
    ratio: "4:3",
    prompt: data.imagePrompt1 || (data.imagePrompt || "Editorial brand event photograph, vibrant, premium quality, no text overlays."),
    google: data.googleSearch1 || (data.googleSearch || "site:unsplash.com brand activation event photography")
  });

  // White gutter between images
  slide.addShape("rect", {
    x: splitX, y: imgHalf, w: imgW, h: gutterH,
    fill: { color: C.white }, line: { color: C.white }
  });

  // Bottom image placeholder
  addImagePlaceholder(slide, splitX, imgHalf + gutterH, imgW, imgHalf, {
    ratio: "4:3",
    prompt: data.imagePrompt2 || (data.imagePrompt || "Editorial brand event photograph, vibrant, premium quality, no text overlays."),
    google: data.googleSearch2 || (data.googleSearch || "site:unsplash.com brand activation event photography")
  });

  addSectionLabel(slide, data.section || "");
  // X white — image sits behind top-right corner
  addXWhite(slide);

  addHeadline(slide, [
    { text: "The ", bold: false },
    { text: data.boldWord || "", bold: true },
    { text: "\n" + (data.rest || ""), bold: false }
  ], { x: 0.29, y: 0.52, w: splitX - 0.4, h: 2.2, size: 28 });

  if (data.body) {
    slide.addText(data.body, {
      x: 0.29, y: 2.85, w: splitX - 0.4, h: 0.85,
      fontSize: 11, fontFace: FONT, color: C.dark,
      valign: "top", margin: 0
    });
  }

  if (data.quote) {
    slide.addText(`"${data.quote}"`, {
      x: 0.29, y: 6.1, w: splitX - 0.4, h: 0.9,
      fontSize: 10, fontFace: FONT, color: C.dark,
      valign: "bottom", italic: true, margin: 0
    });
  }
}

// --- SINGLE IMAGE SLIDE --- //
// Left panel: white with headline + body. Right panel: user-supplied image fills remaining width.
// Gradient strip runs along the bottom of the left panel only.
function addSingleImageSlide(pres, data) {
  const slide = pres.addSlide();
  const splitX = 3.55; // left panel width ~26.6% matching template

  // White left panel
  slide.addShape(pres.ShapeType.rect, {
    x: 0, y: 0, w: splitX, h: H, fill: { color: C.white }, line: { color: C.white }
  });

  // Left panel stays fully white — no gradient patch

  // Right panel: image or grey gradient placeholder with magenta prompts
  if (data.image) {
    slide.addImage({ data: data.image, x: splitX, y: 0, w: W - splitX, h: H,
      sizing: { type: "cover", w: W - splitX, h: H } });
  } else {
    addImagePlaceholder(slide, splitX, 0, W - splitX, H, {
      ratio: data.imageRatio || "4:3",
      prompt: data.imagePrompt || "Premium brand experience or live event photograph. Vibrant, editorial quality, no text overlays.",
      google: data.googleSearch || "site:unsplash.com brand activation event photography"
    });
  }

  addSectionLabel(slide, data.section || "");
  // X always white — image occupies right panel behind it
  addXWhite(slide);

  // Headline
  addHeadline(slide, [
    { text: "The ", bold: false },
    { text: data.boldWord || "", bold: true },
    { text: "\n" + (data.rest || ""), bold: false }
  ], { x: 0.29, y: 0.52, w: splitX - 0.4, h: 2.2, size: 28 });

  // Body paragraphs
  if (data.body) {
    slide.addText(data.body, {
      x: 0.29, y: 2.85, w: splitX - 0.4, h: 0.85,
      fontSize: 11, fontFace: FONT, color: C.dark,
      valign: "top", margin: 0
    });
  }

  // Optional quote at bottom of left panel
  if (data.quote) {
    slide.addText(`"${data.quote}"`, {
      x: 0.29, y: 6.1, w: splitX - 0.4, h: 0.9,
      fontSize: 10, fontFace: FONT, color: C.dark,
      valign: "bottom", italic: true, margin: 0
    });
  }
}

// --- TIMELINE SLIDE --- //
// Two swim-lane timeline matching template page 23.
// Key design decisions:
//   divY      = the horizontal dotted line that separates the two rows
//   topEvents = items above the line — date bold at top, dotted line drops DOWN to divY
//   botEvents = items below the line — date bold just below divY, dotted line drops DOWN from divY
//   Vertical tick lines are drawn intelligently from divY to the text position
function addTimelineSlide(pres, data) {
  const slide = pres.addSlide();

  // Layout constants
  // Equal padding above and below the divider line:
  // topGap = distance from text box bottom to divY
  // botGap = distance from divY to text box top
  // Both are set to the same value so spacing is visually balanced
  const divY     = 3.55;   // horizontal divider line
  const topGap   = 0.45;   // gap between bottom of top text and divider line
  const botGap   = 0.45;   // gap between divider line and top of bottom text (= topGap)
  const topTextH = divY - topGap - 1.3;  // text box height fills from topTextY to line
  const topTextY = 1.3;    // top-row text boxes start here
  const botTextY = divY + botGap; // bottom-row text boxes start here — equal gap below
  const phaseY  = 6.85;   // phase labels row at very bottom
  const leftX   = ML + 1.4;  // start of timeline — leaves room for PRODUCTION/FINANCE labels
  const rightX  = 13.04;  // end of timeline
  const txtW    = 1.25;   // width of each event text box
  const dateFontSize = 8;
  const labelFontSize = 7.5;

  // White upper swim-lane
  slide.addShape(pres.ShapeType.rect, {
    x: 0, y: 0, w: W, h: divY, fill: { color: C.white }, line: { color: C.white }
  });

  // Gradient lower swim-lane
  slide.addImage({ data: gradientWide, x: 0, y: divY, w: W, h: H - divY });

  addSectionLabel(slide, data.section || "");
  addXBlack(slide);

  // Title — bold, top-left
  if (data.title) {
    slide.addText(data.title, {
      x: ML, y: 0.48, w: 5.0, h: 0.7,
      fontSize: 32, fontFace: FONT, bold: true,
      color: C.dark, valign: "top", margin: 0
    });
  }

  // Row labels — rotated 270°, centred vertically in each swim lane, x at ML
  const rowLabels = data.rowLabels || ["PRODUCTION", "FINANCE"];
  const lblSpan = 1.6;  // box width becomes label height after rotation
  slide.addText(rowLabels[0], {
    x: ML - lblSpan / 2 + 0.1, y: divY / 2 - 0.11, w: lblSpan, h: 0.22,
    fontSize: 7, fontFace: FONT, bold: true, color: C.dark,
    align: "center", valign: "middle", rotate: 270, margin: 0
  });
  slide.addText(rowLabels[1], {
    x: ML - lblSpan / 2 + 0.1, y: divY + (H - divY) / 2 - 0.11, w: lblSpan, h: 0.22,
    fontSize: 7, fontFace: FONT, bold: true, color: C.dark,
    align: "center", valign: "middle", rotate: 270, margin: 0
  });

  // Phase vertical dividers and labels
  const phases = data.phases || [];
  phases.forEach((phase, pi) => {
    if (pi > 0) {
      // Black above the gradient divider line
      slide.addShape(pres.ShapeType.line, {
        x: phase.startX, y: 1.45, w: 0, h: divY - 1.45,
        line: { color: C.dark, width: 1.5 }
      });
      // White below
      slide.addShape(pres.ShapeType.line, {
        x: phase.startX, y: divY, w: 0, h: phaseY - divY,
        line: { color: C.white, width: 1.5 }
      });
    }
    slide.addText(phase.label, {
      x: phase.startX, y: phaseY, w: phase.endX - phase.startX, h: 0.38,
      fontSize: 9, fontFace: FONT, bold: true, color: C.dark,
      align: "center", valign: "middle", margin: 0
    });
  });

  // Central horizontal dotted divider
  slide.addShape(pres.ShapeType.line, {
    x: leftX, y: divY, w: rightX - leftX, h: 0,
    line: { color: C.dark, width: 1.0, dashType: "sysDot" }
  });

  // Arrow at right end
  slide.addText("▶", {
    x: rightX + 0.02, y: divY - 0.17, w: 0.3, h: 0.34,
    fontSize: 9, fontFace: FONT, color: C.dark, align: "center", margin: 0
  });

  // Equal padding above and below the divider line:
  // Below gap = botTextY - divY. Above gap mirrors this so bottom of top text box
  // sits the same distance above divY as the top of bottom text box sits below it.
  const belowGap   = botTextY - divY;        // distance from divY to bottom text top
  const topBoxBot  = divY - belowGap;        // where the top text box bottom must sit
  const topBoxH    = topBoxBot - topTextY;   // text box height from topTextY to topBoxBot

  // TOP EVENTS — text box bottom-aligned, equal padding above and below divider line
  const topEvents = data.topEvents || [];
  topEvents.forEach(ev => {
    const tx = ev.x - txtW / 2;

    const evRuns = [
      { text: ev.date,  options: { bold: true,  fontSize: dateFontSize,  fontFace: FONT, color: C.dark } },
      { text: "\n\n",  options: { bold: false, fontSize: dateFontSize,  fontFace: FONT, color: C.dark } },
      { text: ev.label, options: { bold: false, fontSize: labelFontSize, fontFace: FONT, color: C.dark } }
    ];
    slide.addText(evRuns, {
      x: tx, y: topTextY, w: txtW, h: topBoxH,
      align: "center", valign: "bottom", margin: 0
    });

    // Dotted tick line from bottom of text box down to divider dot
    slide.addShape(pres.ShapeType.line, {
      x: ev.x, y: topBoxBot, w: 0, h: divY - topBoxBot,
      line: { color: "999999", width: 0.6, dashType: "sysDot" }
    });

    // Dot on divider line
    slide.addShape(pres.ShapeType.ellipse, {
      x: ev.x - 0.045, y: divY - 0.045, w: 0.09, h: 0.09,
      fill: { color: C.dark }, line: { color: C.dark, width: 0 }
    });
  });

  // BOTTOM EVENTS — text box top-aligned from botTextY, dotted tick from dot down to text top
  const botEvents = data.bottomEvents || [];
  botEvents.forEach(ev => {
    const tx   = ev.x - txtW / 2;
    const boxH = H - botTextY - 0.55;

    // Dotted tick line from divider dot down to top of text box
    slide.addShape(pres.ShapeType.line, {
      x: ev.x, y: divY, w: 0, h: botTextY - divY,
      line: { color: "999999", width: 0.6, dashType: "sysDot" }
    });

    const evRuns = [
      { text: ev.date,  options: { bold: true,  fontSize: dateFontSize,  fontFace: FONT, color: C.dark } },
      { text: "\n\n",  options: { bold: false, fontSize: dateFontSize,  fontFace: FONT, color: C.dark } },
      { text: ev.label, options: { bold: false, fontSize: labelFontSize, fontFace: FONT, color: C.dark } }
    ];
    slide.addText(evRuns, {
      x: tx, y: botTextY, w: txtW, h: boxH,
      align: "center", valign: "top", margin: 0
    });

    // Dot on divider line
    slide.addShape(pres.ShapeType.ellipse, {
      x: ev.x - 0.045, y: divY - 0.045, w: 0.09, h: 0.09,
      fill: { color: C.dark }, line: { color: C.dark, width: 0 }
    });
  });
}

// --- FOUR-COLUMN DETAIL SLIDE --- //
// White top with headline, gradient bottom strip behind four columns
// Each column: bold title, body copy below. Matches template page 15.
function addFourColumnSlide(pres, data) {
  const slide = pres.addSlide();

  // Gradient lower band (bottom ~40% of slide)
  slide.addImage({ data: gradientWide, x: 0, y: 4.5, w: W, h: 3.0 });

  addSectionLabel(slide, data.section || "");
  addXBlack(slide);

  if (data.headline) {
    addHeadline(slide, [
      { text: data.headline.normal || "", bold: true },
      { text: (data.headline.bold ? "\n" + data.headline.bold : ""), bold: false }
    ], { x: ML, y: 0.48, w: 8.0, h: 1.55, size: 32 });
  }

  const cols = data.columns || [];
  const colW = (W - ML - 0.1) / 4;
  const lineTopY = 3.1;   // accent line — title sits above, body below on gradient

  cols.forEach((col, i) => {
    const x  = ML + i * colW;
    const cw = colW - 0.15;

    // Accent line (heavier)
    slide.addShape(pres.ShapeType.line, {
      x, y: lineTopY, w: cw, h: 0,
      line: { color: C.dark, width: 1.5 }
    });

    // Column title ABOVE line — bold, tight box
    slide.addText(col.title, {
      x, y: lineTopY - 0.62, w: cw, h: 0.58,
      fontSize: 13, fontFace: FONT, bold: true,
      color: C.dark, valign: "bottom", margin: 0
    });

    // Body copy below line — sits on gradient
    slide.addText(col.body, {
      x, y: lineTopY + 0.12, w: cw, h: 2.95,
      fontSize: 11, fontFace: FONT, color: C.dark,
      valign: "top", margin: 0
    });
  });
}

// --- KEY STATS SLIDE --- //
// White slide, large stat numbers across the centre with labels below
// Based on template page 35: big number top, label below, thin bottom gradient strip
function addKeyStatsSlide(pres, data) {
  const slide = pres.addSlide();

  slide.addImage({ data: gradientStrip, x: 0, y: 6.92, w: W, h: 0.806 });

  addSectionLabel(slide, data.section || "");
  addXBlack(slide);

  if (data.headline) {
    addHeadline(slide, [
      { text: data.headline.normal || "", bold: true },
      { text: (data.headline.bold ? " " + data.headline.bold : ""), bold: false }
    ], { x: ML, y: 0.48, w: 8.0, h: 1.4, size: 32 });
  }

  const stats = data.stats || [];
  const count = stats.length;
  const colW  = W / count;
  // Vertical centre of stat block — number top, rule, label below
  const numY    = 2.55;
  const numH    = 1.5;
  const ruleY   = numY + numH + 0.08;
  const labelY  = ruleY + 0.12;

  stats.forEach((stat, i) => {
    const cx  = i * colW;
    const pad = colW * 0.12; // inner padding each side

    // Big number — font scales down for longer strings to prevent wrap
    const numLen = String(stat.number).length;
    const numSize = numLen <= 3 ? 72 : numLen <= 5 ? 58 : 46;
    slide.addText(stat.number, {
      x: cx, y: numY, w: colW, h: numH,
      fontSize: numSize, fontFace: FONT, bold: true,
      color: C.dark, align: "center", valign: "middle", margin: 0
    });

    // Thin rule separator (same width as number, centred in column)
    slide.addShape(pres.ShapeType.line, {
      x: cx + pad, y: ruleY, w: colW - pad * 2, h: 0,
      line: { color: C.dark, width: 1.0 }
    });

    // Label — locked to exactly the rule width so it cannot overflow
    // Rule spans cx+pad to cx+colW-pad; label text box matches exactly
    slide.addText(stat.label, {
      x: cx + pad, y: labelY, w: colW - pad * 2, h: 0.65,
      fontSize: 13, fontFace: FONT, color: C.dark,
      align: "center", valign: "top", margin: 0
    });
  });
}

// --- CONTENTS / AGENDA SLIDE --- //
// SECTION LABEL: always "Sections" (not "Contents")
// NUMBERING: sequential ordinals 01, 02, 03 ... (not absolute page numbers)
//            so manually inserted slides do not break the sequence
// LAYOUT: two columns with 1cm gutter — left column fills first
// MAXIMUM: 12 divider sections permitted per deck
//          Layout is tightened to fit 12 items; beyond 12 generator will warn
// PAGE NUMBERS: no page numbers on any slide — removed globally
function addContentsSlide(pres, data) {
  const slide = pres.addSlide();

  // Gradient right panel
  slide.addImage({ data: gradientPanel, x: 8.6, y: 0, w: 4.73, h: H });

  // Section label always reads "Sections"
  addSectionLabel(slide, "Sections");
  addXBlack(slide);

  const items = data.items || [];
  if (items.length > 12) {
    console.warn("CONTENTS: maximum 16 sections exceeded — deck has " + items.length + " dividers. Truncating to 16.");
  }
  const displayItems = items.slice(0, 16);

  // Two-column layout with 1cm (0.39") gutter
  // Left column fills first; right column takes overflow from item 7 onward
  const gutter   = 0.30;   // slightly reduced gutter for more column width
  const colW     = (8.2 - gutter) / 2;   // available width split in two
  const col1X    = 0.29;
  const col2X    = col1X + colW + gutter;
  const startY   = 1.4;
  const rowH     = 0.54;   // tight enough for 12 items, comfortable for fewer
  const lineW    = colW;
  const perCol   = Math.ceil(displayItems.length / 2);

  displayItems.forEach((item, i) => {
    const col   = i < perCol ? 0 : 1;
    const row   = i < perCol ? i : i - perCol;
    const x     = col === 0 ? col1X : col2X;
    const y     = startY + row * rowH;

    // Top border line
    slide.addShape(pres.ShapeType.line, {
      x, y, w: lineW, h: 0,
      line: { color: C.dark, width: 0.5 }
    });

    // Section name with hyperlink
    const labelOpts = {
      x, y: y + 0.06, w: lineW - 0.9, h: rowH - 0.08,
      fontSize: 13, fontFace: FONT, color: C.dark,
      valign: "middle", margin: 0
    };
    if (item.pageRef) labelOpts.hyperlink = { slide: item.pageRef };
    slide.addText(item.label, labelOpts);

    // Sequential ordinal number — 01, 02, 03 ... independent of actual page
    const ordinal = String(i + 1).padStart(2, "0");
    slide.addText(ordinal, {
      x: x + lineW - 0.85, y: y + 0.06, w: 0.85, h: rowH - 0.08,
      fontSize: 13, fontFace: FONT, color: C.dark,
      align: "right", valign: "middle", margin: 0
    });
  });

  // Bottom border after last item in each column
  [0, 1].forEach(col => {
    const x = col === 0 ? col1X : col2X;
    const itemsInCol = col === 0 ? Math.min(perCol, displayItems.length) : Math.max(0, displayItems.length - perCol);
    if (itemsInCol > 0) {
      const y = startY + itemsInCol * rowH;
      slide.addShape(pres.ShapeType.line, {
        x, y, w: lineW, h: 0,
        line: { color: C.dark, width: 0.5 }
      });
    }
  });
}

// --- SUB-DIVIDER SLIDE --- //
// White top half, gradient bottom half
// Title in General Sans bold — bottom of text aligned to top of gradient band (y=3.75)
function addSubDividerSlide(pres, data) {
  const slide = pres.addSlide();
  const splitY = 3.75;

  // White top half
  slide.addShape(pres.ShapeType.rect, {
    x: 0, y: 0, w: W, h: splitY, fill: { color: C.white }, line: { color: C.white }
  });

  // Gradient bottom half
  slide.addImage({ data: gradientHalf, x: 0, y: splitY, w: W, h: H - splitY });

  addSectionLabel(slide, data.section || "");
  addXBlack(slide);

  // Single-line tight box — bottom of box = top of gradient (splitY)
  // Font size 72 at 96dpi ≈ 1.0" tall; box height set to exactly that
  const titleFontH = 1.05; // approximate rendered height of 72pt text in inches
  slide.addText((data.title || "").toUpperCase(), {
    x: 0.29, y: splitY - titleFontH, w: W - 0.58, h: titleFontH,
    fontSize: 72, fontFace: FONT, bold: true,
    color: C.dark, align: "center", valign: "top", margin: 0
  });
}

// --- DIVIDER SLIDE --- //
// Full-bleed gradient, section title centred — General Sans bold, no ghost echo
function addDividerSlide(pres, data) {
  const slide = pres.addSlide();
  slide.addShape(pres.ShapeType.rect, { x: 0, y: 0, w: W, h: H, fill: { color: "BFC8F5" }, line: { color: "BFC8F5" } });
  slide.addImage({ data: gradientFull, x: 0, y: -0.123, w: W, h: 7.746 });
  addXBlack(slide);

  slide.addText((data.title || "").toUpperCase(), {
    x: 0.29, y: 0, w: W - 0.58, h: H,
    fontSize: 72, fontFace: FONT, bold: true,
    color: C.dark, align: "center", valign: "middle", margin: 0
  });
}

// ============================================================
// CG DIAGRAM STYLE GUIDE
// ============================================================
//
// COLOUR PALETTE (from JSX Sales Intelligence engine):
//   Purple:  #9333ea  — strategy, ideas, brand thinking
//   Green:   #10b981  — outcomes, results, success metrics
//   Pink:    #ec4899  — experience, activation, creative
//   Orange:  #f97316  — process, delivery, timelines
//   Blue:    #3b82f6  — data, insight, research
//   CG Blue: #B8C4E8  — default / neutral (brand gradient)
//
// TWO-TIER LOZENGE SYSTEM (mirrors JSX button design language):
//   Tier 1 FILLED:  solid colour fill + white text + borderRadius
//                   → use for: step numbers, primary labels, section headers
//   Tier 2 GHOST:   transparent fill + coloured border (1.5px) + coloured text
//                   → use for: sub-labels, supporting items, callout notes
//   Rule: within one colour family, never two filled lozenges at same level.
//         Filled = declares; Ghost = supports.
//
// DIAGRAM BACKGROUND SYSTEM:
//   White bg slides: use accent colour tints (10% opacity) for card fills
//   Diagram-heavy slides: use JSX gradient bg (purple→blue) with ghosted X
//   Dark diagram slides: white text throughout, filled lozenges in accent colours
//
// STROKE WEIGHTS:
//   Primary lines / circle strokes: 2pt
//   Secondary / connector lines:    1pt  dashed or dotted
//   Ghost lozenge borders:          1.5pt solid
//
// LABEL HIERARCHY:
//   Micro-label:  7–8pt, bold, uppercase, letter-spacing, accent colour
//   Sub-label:    10–11pt, bold, dark or white
//   Body:         11–12pt, regular, dark or white
//   Number:       72pt+, bold — always paired with a rule separator below
//
// GHOSTED X MARK on diagram pages:
//   Placed top-right of diagram area, opacity 6%, accent colour fill
//   Size: ~1.5" × 1.5" — purely decorative, adds brand depth
// ============================================================

// ── DIAGRAM HELPERS ─────────────────────────────────────────────────────────

// Accent colour definitions
const ACCENT = {
  purple: { fill: "C4943A", ghost: "C4943A", tint: "F5EDD6" },
  green:  { fill: "10b981", ghost: "10b981", tint: "ECFDF5" },
  pink:   { fill: "ec4899", ghost: "ec4899", tint: "FDF2F8" },
  orange: { fill: "f97316", ghost: "f97316", tint: "FFF7ED" },
  blue:   { fill: "3b82f6", ghost: "3b82f6", tint: "EFF6FF" },
  cg:     { fill: "B8C4E8", ghost: "7B90D4", tint: "EEF1FF" }
};

// Draw a filled pill (Tier 1) — JSX: solid colour, white text, tight proportions
function addFilledPill(slide, text, x, y, w, h, accentKey, fontSize) {
  const ac = ACCENT[accentKey] || ACCENT.cg;
  slide.addShape("roundRect", {
    x, y, w, h, rectRadius: h / 2,
    fill: { color: ac.fill }, line: { color: ac.fill, width: 0 }
  });
  slide.addText(text, {
    x, y, w, h, fontSize: fontSize || 9, fontFace: FONT,
    bold: false, color: C.white,   // regular not bold — matches JSX fontWeight:600
    align: "center", valign: "middle", margin: 0
  });
}

// Draw a ghost pill (Tier 2) — JSX: 7–8% tint fill, 1px coloured border, coloured text
function addGhostPill(slide, text, x, y, w, h, accentKey, fontSize) {
  const ac = ACCENT[accentKey] || ACCENT.cg;
  // Frosted fill: colour at ~8% opacity (transparency 92)
  slide.addShape("roundRect", {
    x, y, w, h, rectRadius: h / 2,
    fill: { color: ac.fill, transparency: 92 },
    line: { color: ac.ghost, width: 1.0 }   // 1pt not 1.5pt — lighter
  });
  slide.addText(text, {
    x, y, w, h, fontSize: fontSize || 9, fontFace: FONT,
    bold: false, color: ac.ghost,           // regular weight
    align: "center", valign: "middle", margin: 0
  });
}

// Draw ghosted X watermarks — mimics the JSX three-layer pattern
// Large top-right (opacity 5.5%), medium bottom-left (4%), small mid-right (3%)
function addGhostedX(slide, accentKey) {
  // Top-right large — partially off slide edge
  slide.addImage({ data: xIconBlack, x: 10.2, y: -0.8, w: 4.2, h: 4.2,
    transparency: 94 });
  // Bottom-left medium
  slide.addImage({ data: xIconBlack, x: -1.1, y: 4.8, w: 3.2, h: 3.2,
    transparency: 96 });
  // Mid-right small
  slide.addImage({ data: xIconBlack, x: 11.8, y: 2.8, w: 1.8, h: 1.8,
    transparency: 97 });
}

// Build a Google Images search URL from a plain search string
// Strips "site:unsplash.com" prefix if present — Google Images works better without it
function googleImagesUrl(searchString) {
  const cleaned = searchString
    .replace(/^site:[^\s]+\s*/i, '')   // remove site: operator
    .replace(/site:[^\s]+\s*/gi, '')   // remove any other site: operators
    .trim();
  return "https://www.google.com/search?q=" + encodeURIComponent(cleaned) + "&tbm=isch";
}

// Image placeholder helper
// Uses a real image element so PowerPoint right-click → Change Picture works.
// Displays ChatGPT prompt + Google search string in MAGENTA so users cannot miss them.
// RULE: All image slides must use this helper. Never use solid colour shapes as placeholders.
// USAGE BIAS: Collaborate Global is a highly visual company. Use image slides liberally —
//             aim for at least one image slide every 3-4 slides where content permits.
function addImagePlaceholder(slide, x, y, w, h, opts) {
  const ratio  = opts.ratio  || (() => {
    const r = w / h;
    if (r > 1.7)  return "16:9";
    if (r > 1.2)  return "4:3";
    if (r > 0.95) return "1:1";
    if (r > 0.55) return "2:3";
    return "9:16";
  })();

  // Ratio description embedded naturally in the prompt sentence
  const ratioDescMap = {
    "16:9":  "Shot in 16:9 widescreen landscape format.",
    "4:3":   "Shot in 4:3 landscape format.",
    "1:1":   "Shot in 1:1 square format.",
    "2:3":   "Shot in 2:3 portrait format.",
    "9:16":  "Shot in 9:16 tall portrait format."
  };
  const ratioSentence = ratioDescMap[ratio] || ("Shot in " + ratio + " format.");

  // Build final prompt — if caller already embedded a ratio sentence, use as-is;
  // otherwise append the ratio sentence naturally at the end of the prompt.
  const basePrompt = opts.prompt || "High quality editorial photograph, professional lighting, no text overlays.";
  const hasRatio   = /\d+:\d+/.test(basePrompt);
  const prompt     = hasRatio ? basePrompt : (basePrompt.replace(/\.\s*$/, "") + ". " + ratioSentence);
  const google     = opts.google || "site:unsplash.com professional photography";

  // Real image element — right-click → Change Picture in PowerPoint
  slide.addImage({ data: imgPlaceholder, x, y, w, h, sizing: { type: "cover", w, h } });

  // Magenta guidance — must be deleted by user after sourcing the image
  // EC4899 = hot pink/magenta — unmissable, cannot be left in a finished deck
  // ONE text box — one click to select, one delete to remove
  // Google search text is a live hyperlink — opens Google Images with results pre-loaded
  const M   = "EC4899";
  const tx  = x + 0.15;
  const tw  = w - 0.3;
  const gUrl = googleImagesUrl(google);
  slide.addText([
    { text: "ChatGPT image prompt — DELETE AFTER USE:\n", options: { bold: true,  fontSize: 10, fontFace: FONT, color: M } },
    { text: prompt + "\n\n",                             options: { bold: false, fontSize: 10, fontFace: FONT, color: M } },
    { text: "Google Images search — click to open, DELETE AFTER USE:\n", options: { bold: true, fontSize: 10, fontFace: FONT, color: M } },
    { text: google, options: { bold: false, fontSize: 10, fontFace: FONT, color: M, hyperlink: { url: gUrl } } }
  ], { x: tx, y: y + h * 0.28, w: tw, h: h * 0.55, align: "center", valign: "top", margin: 0 });
}

// ── IMAGE & LAYOUT SLIDES ────────────────────────────────────────────────────

// PAGE 24/25 STYLE — left panel white (headline + optional sub + quote),
// right panel image placeholder. Collaborate X floats on top layer.
function addImagePanelSlide(pres, data) {
  const slide = pres.addSlide();
  const splitX = 3.55;
  const hasQuote = !!data.quote;

  // White left panel
  slide.addShape("rect", {
    x: 0, y: 0, w: splitX, h: H, fill: { color: C.white }, line: { color: C.white }
  });

  // Gradient lower portion of left panel
  slide.addImage({ data: gradientPanel, x: 0, y: 3.75, w: splitX, h: H - 3.75 });

  addSectionLabel(slide, data.section || "");

  // Bold headline
  addHeadline(slide, [
    { text: data.boldWord || "", bold: true },
    { text: (data.rest ? "\n" + data.rest : ""), bold: false }
  ], { x: ML, y: 0.52, w: splitX - ML - 0.2, h: 2.2, size: 28 });

  // Optional sub-copy
  if (data.body) {
    slide.addText(data.body, {
      x: ML, y: 2.85, w: splitX - ML - 0.2, h: 0.85,
      fontSize: 11, fontFace: FONT, color: C.dark, valign: "top", margin: 0
    });
  }

  // Optional quote bottom of left panel
  if (hasQuote) {
    slide.addText("\u201C" + data.quote + "\u201D", {
      x: ML, y: 6.1, w: splitX - ML - 0.2, h: 0.9,
      fontSize: 10, fontFace: FONT, color: C.dark,
      italic: true, valign: "bottom", margin: 0
    });
  }

  // Right image panel
  const imgRatio = data.imageRatio || "4:3";
  addImagePlaceholder(slide, splitX, 0, W - splitX, H, {
    ratio: imgRatio,
    label: data.imageLabel || "MAIN IMAGE",
    prompt: (data.imagePrompt || "Premium experiential event or brand activation photograph, dynamic and editorial, no text overlays.") + " Shot in " + imgRatio + " aspect ratio.",
    google: data.googleSearch || "site:unsplash.com brand activation event photography",
    bgColor: "1A1A1A", onGradient: false
  });

  // X icon always white — sits over photograph on right panel
  addXWhite(slide);
}

// PAGE 26/27 STYLE — full-bleed image, gradient overlay bottom third,
// white bold headline + body copy. Three-layer stack.
function addFullBleedTextSlide(pres, data) {
  const slide = pres.addSlide();

  // LAYER 1 — background image (bottom of stack)
  const imgRatio = data.imageRatio || "16:9";
  addImagePlaceholder(slide, 0, 0, W, H, {
    ratio: imgRatio,
    label: data.imageLabel || "FULL BLEED IMAGE",
    prompt: (data.imagePrompt || "Dramatic, cinematic brand event photograph. Dark moody atmosphere, strong lighting, premium feel, no people visible, no text overlays.") + " Shot in " + imgRatio + " aspect ratio, landscape.",
    google: data.googleSearch || "site:unsplash.com cinematic event lighting dramatic",
    bgColor: "1A1A1A"
  });

  // No overlay — white text directly over image
  addSectionLabel(slide, data.section || "", { white: true });
  addXWhite(slide);

  // Bold headline — white, bottom-left
  // All text bold on this slide type per design rules
  const runs = [];
  (data.headline || "").split(" ").forEach((word, i) => {
    if (i > 0) runs.push({ text: " ", options: { fontSize: 48, fontFace: FONT, color: C.white, bold: true } });
    runs.push({
      text: word,
      options: { bold: true, fontSize: 48, fontFace: FONT, color: C.white }
    });
  });
  slide.addText(runs, {
    x: ML, y: 3.2, w: 10.0, h: 3.2,
    valign: "bottom", margin: 0
  });

  // Body copy — white, with increased padding (0.42") below headline for breathing room
  if (data.body) {
    slide.addText(data.body, {
      x: ML, y: 6.65, w: 9.0, h: 0.55,
      fontSize: 13, fontFace: FONT, bold: true, color: C.white,
      valign: "top", margin: 0
    });
  }
}

// PAGE 28 STYLE — four-image grid
function addFourImageGridSlide(pres, data) {
  const slide = pres.addSlide();
  addSectionLabel(slide, data.section || "");

  if (data.headline) {
    slide.addText(data.headline, {
      x: ML, y: 0.45, w: 8.0, h: 0.7,
      fontSize: 24, fontFace: FONT, bold: true,
      color: C.dark, valign: "top", margin: 0
    });
  }

  const images = data.images || [{},{},{},{}];
  const startY = data.headline ? 1.3 : 0.4;
  const gridH  = H - startY - 0.1;
  const cellW  = (W - 0.06) / 2;
  const cellH  = (gridH - 0.06) / 2;

  [[0,0],[1,0],[0,1],[1,1]].forEach(([col, row], i) => {
    const img = images[i] || {};
    const cx  = col * (cellW + 0.06);
    const cy  = startY + row * (cellH + 0.06);
    addImagePlaceholder(slide, cx, cy, cellW, cellH, {
      ratio: img.ratio || "4:3",
      prompt: (img.prompt || "Portfolio shot — brand activation, experiential event or creative installation. Vibrant, editorial quality, no text overlays.") + " Shot in " + (img.ratio || "4:3") + " aspect ratio.",
      google: img.google || "site:unsplash.com brand activation photography"
    });
  });
  // X white — z-ordered above all images
  addXWhite(slide);
}

// PAGE 29 STYLE — mood board (5-cell asymmetric grid)
function addMoodBoardSlide(pres, data) {
  const slide = pres.addSlide();
  // Asymmetric layout: 1 large left, 2 stacked right top, 2 stacked right bottom
  // [x, y, w, h, ratio]
  const cells = [
    [0,      0,    6.0,  H,    "2:3",  "HERO IMAGE"],
    [6.06,   0,    3.6,  3.65, "4:3",  "IMAGE 2"],
    [9.72,   0,    3.61, 3.65, "4:3",  "IMAGE 3"],
    [6.06,   3.71, 3.6,  3.79, "4:3",  "IMAGE 4"],
    [9.72,   3.71, 3.61, 3.79, "4:3",  "IMAGE 5"]
  ];
  const images = data.images || [];

  cells.forEach(([x, y, w, h, ratio, defLabel], i) => {
    const img = images[i] || {};
    addImagePlaceholder(slide, x, y, w, h, {
      ratio: img.ratio || ratio,
      prompt: (img.prompt || "Mood board editorial photograph for a premium brand experience. Creative, atmospheric, high quality.") + " Shot in " + (img.ratio || ratio) + " aspect ratio.",
      google: img.google || "site:unsplash.com editorial brand mood"
    });
  });
  // X white — z-ordered above all images
  addXWhite(slide);
}

// PAGE 33/34 STYLE — case study layout
// Left: section label, bold headline, challenge/idea/experience fields on white
// Right: 2×2 grid of case study images
function addCaseStudySlide(pres, data) {
  const slide = pres.addSlide();
  const splitX = 4.0;

  // White left panel
  slide.addShape("rect", {
    x: 0, y: 0, w: splitX, h: H, fill: { color: C.white }, line: { color: C.white }
  });

  addSectionLabel(slide, data.section || "Case Study");
  // X drawn last so it sits above the image grid on the right
  // (called again at end of function after images are drawn)

  // Bold headline
  addHeadline(slide, [
    { text: data.boldWord || "", bold: true },
    { text: (data.rest ? "\n" + data.rest : ""), bold: false }
  ], { x: ML, y: 0.52, w: splitX - ML - 0.2, h: 1.8, size: 26 });

  // Labelled sections
  const sections = data.sections || [];
  let yPos = 2.45;
  sections.forEach(sec => {
    slide.addText(sec.label + ":", {
      x: ML, y: yPos, w: splitX - ML - 0.2, h: 0.25,
      fontSize: 11, fontFace: FONT, bold: true, color: C.dark, valign: "top", margin: 0
    });
    slide.addText(sec.body, {
      x: ML, y: yPos + 0.27, w: splitX - ML - 0.2, h: 0.9,
      fontSize: 10, fontFace: FONT, color: C.dark, valign: "top", margin: 0
    });
    yPos += 1.25;
  });

  // Optional italic italic italic italic italic italic conclusion
  if (data.conclusion) {
    slide.addText(data.conclusion, {
      x: ML, y: yPos + 0.1, w: splitX - ML - 0.2, h: 1.0,
      fontSize: 10, fontFace: FONT, color: C.dark, italic: true, valign: "top", margin: 0
    });
  }

  // Right side: 2×2 image grid
  const images = data.images || [{},{},{},{}];
  const imgStartX = splitX + 0.06;
  const imgW = (W - imgStartX - 0.06) / 2;
  const imgH = (H - 0.06) / 2;

  [[0,0],[1,0],[0,1],[1,1]].forEach(([col, row], i) => {
    const img = images[i] || {};
    const cx  = imgStartX + col * (imgW + 0.06);
    const cy  = row * (imgH + 0.06);
    addImagePlaceholder(slide, cx, cy, imgW, imgH, {
      ratio: img.ratio || "4:3",
      prompt: (img.prompt || "Case study event photography — behind the scenes, live moments, brand activation. Premium, editorial quality, no text overlays.") + " Shot in " + (img.ratio || "4:3") + " aspect ratio.",
      google: img.google || "site:unsplash.com brand event case study"
    });
  });
  // X white — z-ordered above image grid
  addXWhite(slide);
}

// PAGE 35 STYLE — case study hero full-bleed with overlay
// Same three-layer stack as page 26 but with stat callouts
function addCaseStudyHeroSlide(pres, data) {
  const slide = pres.addSlide();

  // LAYER 1 — background image
  addImagePlaceholder(slide, 0, 0, W, H, {
    ratio: "16:9",
    label: "CASE STUDY HERO",
    prompt: (data.imagePrompt || "Dramatic hero shot of a live brand activation or event. Wide angle, cinematic, premium, no text overlays.") + " Shot in 16:9 aspect ratio, landscape.",
    google: data.googleSearch || "site:unsplash.com brand activation hero shot wide angle cinematic",
    bgColor: "111111"
  });

  // No overlay — white text directly over image
  addXWhite(slide);

  // CLIENT LOGO placeholder removed — not used in standard deck output

  // Bold headline bottom-left
  const runs = [];
  (data.headline || "").split(" ").forEach((word, i) => {
    if (i > 0) runs.push({ text: " ", options: { fontSize: 48, fontFace: FONT, color: C.white } });
    runs.push({
      text: word,
      options: { bold: word === (data.boldWord || ""), fontSize: 48, fontFace: FONT, color: C.white }
    });
  });
  slide.addText(runs, {
    x: ML, y: 3.8, w: 9.0, h: 2.8, valign: "bottom", margin: 0
  });

  // Stat callouts — right side, stacked with dividers
  const stats = data.stats || [];
  stats.forEach((stat, i) => {
    const sy = 4.5 + i * 0.9;
    if (i > 0) {
      slide.addShape("line", {
        x: 10.5, y: sy - 0.08, w: 2.65, h: 0,
        line: { color: C.white, width: 0.5 }
      });
    }
    slide.addText(stat, {
      x: 10.5, y: sy, w: 2.65, h: 0.8,
      fontSize: 13, fontFace: FONT, color: C.white,
      align: "center", valign: "middle", margin: 0
    });
  });
}

// DEVICE MOCKUP SLIDES ───────────────────────────────────────────────────────

// Draw an iPhone frame (vector, no image needed for the chrome)
// screenX/Y/W/H = the zone where the user's image goes
function drawIPhoneFrame(slide, x, y, w, h) {
  const r     = w * 0.12;  // corner radius proportional to width
  const bezel = w * 0.055; // bezel width
  const notchW = w * 0.35;
  const notchH = h * 0.025;
  const notchX = x + (w - notchW) / 2;
  const notchY = y + bezel * 0.7;

  // Outer frame
  slide.addShape("roundRect", {
    x, y, w, h, rectRadius: r,
    fill: { color: "1A1A1A" }, line: { color: "444444", width: 1.5 }
  });

  // Screen area (inner rect)
  const sx = x + bezel;
  const sy = y + bezel * 1.8;
  const sw = w - bezel * 2;
  const sh = h - bezel * 3.2;
  slide.addShape("rect", {
    x: sx, y: sy, w: sw, h: sh,
    fill: { color: "2A2A2A" }, line: { color: "333333", width: 0.5 }
  });

  // Notch
  slide.addShape("roundRect", {
    x: notchX, y: notchY, w: notchW, h: notchH,
    rectRadius: notchH / 2,
    fill: { color: "1A1A1A" }, line: { color: "1A1A1A", width: 0 }
  });

  // Side buttons (volume)
  slide.addShape("rect", {
    x: x - 0.04, y: y + h * 0.28, w: 0.04, h: h * 0.08,
    fill: { color: "333333" }, line: { color: "333333", width: 0 }
  });
  slide.addShape("rect", {
    x: x - 0.04, y: y + h * 0.38, w: 0.04, h: h * 0.08,
    fill: { color: "333333" }, line: { color: "333333", width: 0 }
  });

  // Power button
  slide.addShape("rect", {
    x: x + w, y: y + h * 0.32, w: 0.04, h: h * 0.1,
    fill: { color: "333333" }, line: { color: "333333", width: 0 }
  });

  return { x: sx, y: sy, w: sw, h: sh }; // return screen zone
}

// PAGE 30/31 STYLE — hand holding iPhone with real image placeholder on screen
// Uses the extracted phone image from the template PDF.
// Screen placeholder is a real image element for right-click → Change Picture.
function addDeviceMockupSlide(pres, data) {
  const slide = pres.addSlide();
  const splitX = 3.55;

  // White left panel
  slide.addShape("rect", {
    x: 0, y: 0, w: splitX, h: H, fill: { color: C.white }, line: { color: C.white }
  });

  // Left panel stays fully white

  addSectionLabel(slide, data.section || "");
  addXBlack(slide);

  addHeadline(slide, [
    { text: "The ", bold: false },
    { text: data.boldWord || "", bold: true },
    { text: "\n" + (data.rest || ""), bold: false }
  ], { x: ML, y: 0.52, w: splitX - ML - 0.2, h: 2.2, size: 28 });

  if (data.body) {
    slide.addText(data.body, {
      x: ML, y: 2.85, w: splitX - ML - 0.2, h: 0.75,
      fontSize: 11, fontFace: FONT, color: C.dark, valign: "top", margin: 0
    });
  }

  // Vector iPhone frame centred in right panel
  const rightW  = W - splitX;
  const dw      = rightW * 0.52;
  const dh      = dw * (19.5 / 9);
  const dx      = splitX + (rightW - dw) / 2;
  const dy      = (H - dh) / 2;

  // Gradient background behind phone
  slide.addImage({ data: gradientPanel, x: splitX, y: 0, w: rightW, h: H });

  // Draw vector iPhone frame — returns screen zone coordinates
  const screen = drawIPhoneFrame(slide, dx, dy, dw, dh);

  // Screen placeholder — magenta image prompt so user knows to replace
  const screenPrompt = data.imagePrompt || "Mobile app UI or social content screenshot. Portrait 9:16, no device frame, clean and branded.";
  const screenGoogle = data.googleSearch || "site:unsplash.com mobile app UI screenshot portrait";

  slide.addImage({ data: imgPlaceholder, x: screen.x, y: screen.y, w: screen.w, h: screen.h,
    sizing: { type: "cover", w: screen.w, h: screen.h } });

  // ONE text box — one click to delete
  const M  = "EC4899";
  const sx2 = screen.x + 0.05;
  const sw2 = screen.w - 0.1;
  slide.addText([
    { text: "ChatGPT prompt — DELETE AFTER USE:\n", options: { bold: true,  fontSize: 7, fontFace: FONT, color: M } },
    { text: screenPrompt + "\n\n",                 options: { bold: false, fontSize: 7, fontFace: FONT, color: M } },
    { text: "Google search — DELETE AFTER USE:\n",  options: { bold: true,  fontSize: 7, fontFace: FONT, color: M } },
    { text: screenGoogle,                            options: { bold: false, fontSize: 7, fontFace: FONT, color: M } }
  ], { x: sx2, y: screen.y + screen.h * 0.22, w: sw2, h: screen.h * 0.6,
       align: "center", valign: "top", margin: 0 });
}

// PAGE 32 STYLE — three phones on white background, no tile borders
// Three hand-holding-phone images at varying x positions and scales
// Screen placeholder on each phone as real image element
function addDeviceMoodBoardSlide(pres, data) {
  const slide = pres.addSlide();

  // White background
  addXBlack(slide);

  // Three vector iPhone frames at varying scales, centred across slide
  // Layout: left phone slightly smaller, centre full-height, right slightly smaller
  const phoneConfigs = [
    { cx: 2.8,   heightScale: 0.80 },
    { cx: 6.665, heightScale: 0.92 },
    { cx: 10.53, heightScale: 0.80 }
  ];
  const images = data.images || [];

  phoneConfigs.forEach((cfg, i) => {
    const ph = H * cfg.heightScale;
    const pw = ph * (9 / 19.5);  // correct iPhone aspect ratio
    const px = cfg.cx - pw / 2;
    const py = (H - ph) / 2;

    const screen = drawIPhoneFrame(slide, px, py, pw, ph);

    const img = images[i] || {};
    const prompt = (img.prompt || "Mobile app screenshot or digital content, portrait, no device chrome.") + " 9:16 portrait ratio.";
    console.log("\n[IMAGE PLACEHOLDER] SCREEN " + (i+1) + " (9:16)");
    console.log('  ChatGPT: "' + prompt + '"');
    console.log("  Google:  " + (img.google || "mobile app UI screenshot portrait") + "\n");

    slide.addImage({ data: imgPlaceholder, x: screen.x, y: screen.y, w: screen.w, h: screen.h,
      sizing: { type: "cover", w: screen.w, h: screen.h } });

    // ONE text box — one click to delete
    const M2 = "EC4899";
    const bprompt = img.prompt || "Mobile content screenshot — app UI, social feed or branded digital experience. Portrait 9:16 format.";
    slide.addText([
      { text: "ChatGPT prompt — DELETE AFTER USE:\n", options: { bold: true,  fontSize: 6.5, fontFace: FONT, color: M2 } },
      { text: bprompt,                                 options: { bold: false, fontSize: 6.5, fontFace: FONT, color: M2 } }
    ], { x: screen.x + 0.04, y: screen.y + screen.h * 0.25,
         w: screen.w - 0.08, h: screen.h * 0.55,
         align: "center", valign: "top", margin: 0 });
  });
}

// ── DIAGRAM SLIDES ───────────────────────────────────────────────────────────

// CONCENTRIC CIRCLES — page 9 of PDF
// Three nested rings, each with title + body in the ring, leader lines to left notes
function addConcentricSlide(pres, data) {
  const slide = pres.addSlide();

  // Diagram-heavy slide: JSX gradient background (purple→blue tint)
  // Left third: lightest pink tint from JSX palette
  slide.addShape("rect", {
    x: 0, y: 0, w: W / 3, h: H,
    fill: { color: "FDF4FF" }, line: { color: "FDF4FF" }
  });
  // Right two-thirds: clean gradient (no diagonal artefact)
  slide.addImage({ data: gradientFull, x: W / 3, y: 0, w: W - W / 3, h: H });

  addSectionLabel(slide, data.section || "");
  addXBlack(slide);

  // Ghosted X watermark on diagram side
  addGhostedX(slide, "purple");

  // Headline left panel
  const leftW = W / 3;
  addHeadline(slide, [
    { text: "The ", bold: false },
    { text: data.boldWord || "", bold: true },
    { text: "\n" + (data.rest || ""), bold: false }
  ], { x: ML, y: 0.52, w: leftW - ML - 0.2, h: 2.0, size: 26 });

  // Three concentric circles — scaled 20%
  // Pills at TOP of each ring (moved down 10px from ring edge for breathing room)
  // Body text at BOTTOM of each band
  // Leader lines from notes precisely touch the left edge of each respective ring
  const circles    = data.circles || [];
  const radii      = [3.18, 2.22, 1.26];
  const accentKeys = ["purple", "pink", "blue"];
  const cx = W / 3 + (W - W / 3) / 2 + 0.05;
  const cy = H / 2 + 0.15;

  // Precompute ring left-edge x positions for precise leader line termini
  // cx - r gives the geometric left edge of each circle
  // We add a tiny offset (+0.02) so the line visually touches rather than stops short
  const ringLeftEdges = radii.map(r => cx - r + 0.02);

  // Notes + leader lines — drawn before circles so lines sit behind rings
  const notes = data.notes || [];
  const noteAccents = ["purple", "pink", "blue"];
  notes.forEach((note, i) => {
    const ny       = 2.8 + i * 0.95;
    const lineY    = ny + 0.11;
    const lineEndX = ringLeftEdges[i];

    addGhostPill(slide, "Note " + (i+1), ML, ny, 0.68, 0.22, noteAccents[i], 8);
    slide.addText(note, {
      x: ML + 0.74, y: ny, w: leftW - ML - 0.9, h: 0.22,
      fontSize: 10, fontFace: FONT, color: C.dark, valign: "middle", margin: 0
    });

    // Leader line from left panel edge precisely to the left edge of this ring
    slide.addShape("line", {
      x: leftW - 0.05, y: lineY, w: lineEndX - (leftW - 0.05), h: 0,
      line: { color: ACCENT[noteAccents[i]].ghost, width: 0.75, dashType: "dash" }
    });
  });

  // Draw circles back-to-front (outermost first)
  circles.slice(0, 3).forEach((circ, i) => {
    const r  = radii[i];
    const ac = ACCENT[accentKeys[i]];
    slide.addShape("ellipse", {
      x: cx - r, y: cy - r, w: r * 2, h: r * 2,
      fill: { color: ac.fill, transparency: i === 0 ? 88 : i === 1 ? 84 : 80 },
      line: { color: ac.fill, width: 0.75 }
    });
  });

  // Labels on top of circles
  circles.slice(0, 3).forEach((circ, i) => {
    const r     = radii[i];
    const rNext = i < 2 ? radii[i + 1] : 0;
    const pillW = Math.min(r * 1.1, 1.6);
    const pillH = 0.30;

    // Pill: top of ring + 0.28" (was 0.14") — ~10px more breathing room from ring edge
    addFilledPill(slide, circ.title || "", cx - pillW / 2, cy - r + 0.28, pillW, pillH, accentKeys[i], 10);

    // Body text in band between this ring and next inner ring
    if (circ.body) {
      const bodyBandTop = cy + rNext + 0.12;
      const bodyBandBot = cy + r - 0.16;
      const bodyH       = Math.max(bodyBandBot - bodyBandTop, 0.35);

      // For inner circle (i===2): centre body text vertically in ring, shifted up slightly
      const innerOffsetY = i === 2 ? -0.18 : 0;

      slide.addText(circ.body, {
        x: cx - r + 0.3, y: bodyBandTop + innerOffsetY, w: (r * 2) - 0.6, h: bodyH,
        fontSize: 9, fontFace: FONT, color: C.dark,
        align: "center", valign: i === 2 ? "middle" : "top", margin: 0
      });
    }
  });
}

// VENN DIAGRAM — page 10 of PDF
// Three overlapping circles, outer dashed containing ring, labels in each zone
function addVennSlide(pres, data) {
  const slide = pres.addSlide();

  // Left third: lightest pink tint from JSX palette
  slide.addShape("rect", {
    x: 0, y: 0, w: W / 3, h: H,
    fill: { color: "FDF4FF" }, line: { color: "FDF4FF" }
  });
  // Right two-thirds: clean gradient (no diagonal artefact)
  slide.addImage({ data: gradientFull, x: W / 3, y: 0, w: W - W / 3, h: H });

  addSectionLabel(slide, data.section || "");
  addXBlack(slide);
  addGhostedX(slide, "blue");

  const leftWv = W / 3;
  addHeadline(slide, [
    { text: "The ", bold: false },
    { text: data.boldWord || "", bold: true },
    { text: "\n" + (data.rest || ""), bold: false }
  ], { x: ML, y: 0.52, w: leftWv - ML - 0.2, h: 2.0, size: 26 });

  const subs = data.subs || [];
  subs.forEach((sub, i) => {
    slide.addText([
      { text: sub.label + ": ", options: { bold: true, fontSize: 11, fontFace: FONT, color: C.dark } },
      { text: sub.body, options: { bold: false, fontSize: 11, fontFace: FONT, color: C.dark } }
    ], { x: ML, y: 2.75 + i * 1.1, w: leftWv - ML - 0.2, h: 1.0, valign: "top", margin: 0 });
  });

  // Three overlapping circles — scaled up 20%, labels repositioned into non-intersecting zones
  // Brand = top circle (label moved up), Audience = bottom-left (label moved left),
  // Craft = bottom-right (label moved right). Outer ring enlarged to fully contain all three.
  const circles = data.circles || [{label:"Thing 1"},{label:"Thing 2"},{label:"Thing 3"}];
  const accs    = ["purple", "green", "pink"];
  const r       = 1.74;  // 1.45 * 1.2 = 20% larger
  const offset  = r * 0.85;
  const cx      = W / 3 + (W - W / 3) / 2 + 0.1;
  const cy      = H / 2 + 0.05;

  // Circle centres — equilateral triangle layout
  const centres = [
    { x: cx,                   y: cy - offset * 0.75 }, // top (Brand)
    { x: cx - offset * 0.87,  y: cy + offset * 0.5  }, // bottom-left (Audience)
    { x: cx + offset * 0.87,  y: cy + offset * 0.5  }  // bottom-right (Craft)
  ];

  // Label offset vectors — push each label into the non-intersecting area of its circle
  const labelOffsets = [
    { dx: 0,     dy: -0.55 }, // Brand: up
    { dx: -0.55, dy:  0.25 }, // Audience: left
    { dx:  0.55, dy:  0.25 }  // Craft: right
  ];

  circles.slice(0, 3).forEach((circ, i) => {
    const cc  = centres[i];
    const ac  = ACCENT[accs[i]];
    const lof = labelOffsets[i];

    slide.addShape("ellipse", {
      x: cc.x - r, y: cc.y - r, w: r * 2, h: r * 2,
      fill: { color: ac.fill, transparency: 75 },
      line: { color: ac.fill, width: 0.75 }
    });

    // Label — positioned in non-intersecting arc of each circle
    slide.addText(circ.label, {
      x: cc.x + lof.dx - 0.85, y: cc.y + lof.dy - 0.22, w: 1.7, h: 0.36,
      fontSize: 13, fontFace: FONT, bold: true, color: C.dark,
      align: "center", valign: "middle", margin: 0
    });
    if (circ.body) {
      // Body text kept short — max ~15 words to fit in the arc
      slide.addText(circ.body, {
        x: cc.x + lof.dx - 0.82, y: cc.y + lof.dy + 0.16, w: 1.64, h: 0.65,
        fontSize: 8.5, fontFace: FONT, color: C.dark,
        align: "center", valign: "top", margin: 0
      });
    }
  });

  // Geometric centre of the three circles
  // Top circle centred at cy - offset*0.75, bottom two at cy + offset*0.5
  // True centroid y = average of the three centres
  const topCY    = cy - offset * 0.75;
  const botCY    = cy + offset * 0.5;
  const geoCY    = (topCY + botCY + botCY) / 3;  // weighted — two bottom, one top

  // Outer ring centred on geometric centre, sized to fully contain all circles
  // Furthest point = bottom circles centre + r, so outerR must reach from geoCY
  const outerR   = offset * 0.87 + r + 0.22;
  const ringCY   = geoCY;  // shift ring down to match true centre of three circles

  slide.addShape("ellipse", {
    x: cx - outerR, y: ringCY - outerR, w: outerR * 2, h: outerR * 2,
    fill: { color: C.white, transparency: 100 },
    line: { color: C.dark, width: 1.0, dashType: "dash" }
  });

  // Outer ring label — positioned above the ring
  if (data.outerLabel) {
    addGhostPill(slide, data.outerLabel, cx - 1.1, ringCY - outerR - 0.30, 2.2, 0.24, "cg", 9);
  }
}

// CONVERGENCE DIAGRAM — page 17 of PDF
// Ranked items (#1 #2 #3) with bold value, curly brace below, converging to headline + body
function addConvergenceSlide(pres, data) {
  const slide = pres.addSlide();

  slide.addImage({ data: gradientStrip, x: 0, y: 6.92, w: W, h: 0.806 });

  addSectionLabel(slide, data.section || "");
  addXBlack(slide);

  addHeadline(slide, [
    { text: "The ", bold: false },
    { text: data.boldWord || "", bold: true },
    { text: "\n" + (data.rest || ""), bold: false }
  ], { x: ML, y: 0.48, w: 6.0, h: 1.5, size: 32 });

  const items = data.items || [];
  const accs  = ["purple", "orange", "green"];
  const count = Math.min(items.length, 3);
  const itemW = (W - ML * 2) / Math.max(count, 1);

  items.slice(0, 3).forEach((item, i) => {
    const ix  = ML + i * itemW + itemW / 2;
    const ac  = ACCENT[accs[i]];
    const pillW = Math.min(itemW * 0.72, 1.8);

    // Rank label — ghost pill (frosted, not filled)
    addGhostPill(slide, "#" + (i+1) + "  " + (item.rank || ""), ix - pillW/2, 2.25, pillW, 0.28, accs[i], 9);

    // Large architectural value — typography does the work
    slide.addText(item.value || "", {
      x: ix - itemW * 0.4, y: 2.65, w: itemW * 0.8, h: 0.55,
      fontSize: 22, fontFace: FONT, bold: true, color: ac.fill,
      align: "center", valign: "middle", margin: 0
    });

    // Descriptor — light, regular
    slide.addText(item.label || "", {
      x: ix - itemW * 0.4, y: 3.22, w: itemW * 0.8, h: 0.3,
      fontSize: 12, fontFace: FONT, bold: false, color: C.dark,
      align: "center", valign: "middle", margin: 0
    });
  });

  // Curly brace — dynamically drawn to span full width of all ranked items
  // Built from line segments to precisely mirror a downward-pointing } brace.
  // Spans from the left edge of item 1 to the right edge of the last item.
  // Converges to a centre point pointing down to the outcome pill below.
  const count_i = items.length;
  const itemSpan = (W - ML * 2) / Math.max(count_i, 1);
  const braceL   = ML;                        // left edge of first item
  const braceR   = ML + count_i * itemSpan;   // right edge of last item
  const braceM   = (braceL + braceR) / 2;    // centre point
  const braceY   = 3.6;                       // vertical position of brace
  const dipD     = 0.18;                      // depth of the dip from arms to centre tip
  const armLen   = (braceR - braceL) / 2 - 0.15; // half-span minus centre gap
  const col      = "AAAAAA";
  const lw       = 1.2;

  // Left arm: horizontal from braceL to centre-left, then diagonal dip to centre tip
  // Right arm: mirrors left arm
  // Pattern: outer horizontal → short diagonal down → inner horizontal → centre point

  // Each arm: outer horizontal → knee drop → inner horizontal meeting exactly at braceM
  // Left arm
  const kneeL = braceL + armLen * 0.45;   // left knee x position
  slide.addShape("line", { x: braceL, y: braceY, w: kneeL - braceL, h: 0,
    line: { color: col, width: lw } });   // left outer horizontal
  slide.addShape("line", { x: kneeL, y: braceY, w: 0, h: dipD,
    line: { color: col, width: lw } });   // left knee drop
  slide.addShape("line", { x: kneeL, y: braceY + dipD, w: braceM - kneeL, h: 0,
    line: { color: col, width: lw } });   // left inner horizontal → exactly braceM

  // Right arm (mirror of left)
  const kneeR = braceR - armLen * 0.45;   // right knee x position
  slide.addShape("line", { x: kneeR, y: braceY, w: braceR - kneeR, h: 0,
    line: { color: col, width: lw } });   // right outer horizontal
  slide.addShape("line", { x: kneeR, y: braceY, w: 0, h: dipD,
    line: { color: col, width: lw } });   // right knee drop
  slide.addShape("line", { x: braceM, y: braceY + dipD, w: kneeR - braceM, h: 0,
    line: { color: col, width: lw } });   // right inner horizontal ← exactly braceM

  // Centre vertical drop — starts exactly at braceM, same y as inner horizontals
  slide.addShape("line", { x: braceM, y: braceY + dipD, w: 0, h: 0.28,
    line: { color: col, width: lw } });

  // Convergence outcome below
  if (data.convergence) {
    addGhostPill(slide, data.convergence.title, braceM - 1.4, braceY + dipD + 0.36, 2.8, 0.34, "purple", 13);
    slide.addText(data.convergence.body || "", {
      x: braceM - 2.5, y: braceY + dipD + 0.80, w: 5.0, h: 0.7,
      fontSize: 12, fontFace: FONT, color: C.dark,
      align: "center", valign: "top", margin: 0
    });
  }
}

// PROCESS FLOW — page 18 of PDF
// Horizontal numbered milestone chain with phase labels
function addProcessFlowSlide(pres, data) {
  const slide = pres.addSlide();

  slide.addImage({ data: gradientStrip, x: 0, y: 6.92, w: W, h: 0.806 });

  addSectionLabel(slide, data.section || "");
  addXBlack(slide);

  addHeadline(slide, [
    { text: data.boldWord || "", bold: true },
    { text: (data.rest ? "\n" + data.rest : ""), bold: false }
  ], { x: ML, y: 0.48, w: 6.0, h: 1.4, size: 32 });

  const steps = data.steps || [];
  const count = steps.length;
  const accs  = ["purple","blue","green","orange","pink","cg"];
  const lineY = 3.6; // horizontal spine
  const stepW = (W - ML * 2) / Math.max(count, 1);

  // Horizontal spine line
  slide.addShape("line", {
    x: ML, y: lineY, w: W - ML * 2, h: 0,
    line: { color: "CCCCCC", width: 1.5 }
  });

  steps.forEach((step, i) => {
    const sx  = ML + i * stepW + stepW / 2;
    const acc = accs[i % accs.length];
    const ac  = ACCENT[acc];

    // Node circle on spine
    slide.addShape("ellipse", {
      x: sx - 0.2, y: lineY - 0.2, w: 0.4, h: 0.4,
      fill: { color: ac.fill }, line: { color: ac.fill, width: 0 }
    });
    // Number in node
    slide.addText(String(i+1), {
      x: sx - 0.2, y: lineY - 0.2, w: 0.4, h: 0.4,
      fontSize: 9, fontFace: FONT, bold: true, color: C.white,
      align: "center", valign: "middle", margin: 0
    });

    // Label above spine
    slide.addText(step.date || "", {
      x: sx - stepW * 0.4, y: lineY - 1.0, w: stepW * 0.8, h: 0.28,
      fontSize: 8, fontFace: FONT, bold: true, color: C.dark,
      align: "center", valign: "middle", margin: 0
    });
    slide.addText(step.label || "", {
      x: sx - stepW * 0.4, y: lineY - 0.7, w: stepW * 0.8, h: 0.65,
      fontSize: 9, fontFace: FONT, color: C.dark,
      align: "center", valign: "top", margin: 0
    });

    // Sub-label below spine — ghost pill
    if (step.sub) {
      addGhostPill(slide, step.sub, sx - 0.75, lineY + 0.32, 1.5, 0.26, acc, 8);
    }

    // Phase label bottom
    if (step.phase) {
      slide.addText(step.phase, {
        x: sx - stepW * 0.4, y: lineY + 0.72, w: stepW * 0.8, h: 0.4,
        fontSize: 9, fontFace: FONT, bold: true, color: ac.ghost,
        align: "center", valign: "middle", margin: 0
      });
    }

    // Arrow to next
    if (i < count - 1) {
      slide.addText("›", {
        x: sx + stepW * 0.38, y: lineY - 0.15, w: 0.2, h: 0.3,
        fontSize: 16, fontFace: FONT, color: "CCCCCC",
        align: "center", margin: 0
      });
    }
  });
}

// STRATEGY PILLARS — page 20 of PDF
// Central platform statement, supporting pillars in accent colours
function addStrategyPillarsSlide(pres, data) {
  const slide = pres.addSlide();

  // JSX-style gradient background for diagram pages
  slide.addShape("rect", {
    x: 0, y: 0, w: W, h: H, fill: { color: "FAF8FF" }, line: { color: "FAF8FF" }
  });
  // Right panel — use gradientFull to avoid diagonal bar artefact
  slide.addImage({ data: gradientFull, x: 8.5, y: 0, w: 4.83, h: H });

  addSectionLabel(slide, data.section || "");
  addXBlack(slide);
  addGhostedX(slide, "purple");

  addHeadline(slide, [
    { text: data.boldWord || "", bold: true },
    { text: (data.rest ? "\n" + data.rest : ""), bold: false }
  ], { x: ML, y: 0.48, w: 7.0, h: 1.5, size: 32 });

  // Pillars — computed first so platform lozenge can span their full width
  const pillars = data.pillars || [];
  const accs    = ["purple","green","pink","orange","blue","cg"];
  const pillarW = Math.min((W - ML * 2) / Math.max(pillars.length,1), 2.2);
  const totalPillarW = pillars.length * pillarW + Math.max(pillars.length - 1, 0) * 0.15;

  // Central platform lozenge — spans exactly from left edge of first pillar to right edge of last
  // Width is dynamically computed so it always matches regardless of pillar count (2-5)
  if (data.platform) {
    slide.addShape("roundRect", {
      x: ML, y: 2.1, w: totalPillarW, h: 0.58, rectRadius: 0.29,
      fill: { color: ACCENT.purple.fill, transparency: 88 },
      line: { color: ACCENT.purple.fill, width: 1.0 }
    });
    slide.addText(data.platform, {
      x: ML, y: 2.1, w: totalPillarW, h: 0.58,
      fontSize: 14, fontFace: FONT, bold: false, color: ACCENT.purple.fill,
      align: "center", valign: "middle", margin: 0
    });
  }

  pillars.forEach((pillar, i) => {
    const px  = ML + i * (pillarW + 0.15);
    const acc = accs[i % accs.length];
    const ac  = ACCENT[acc];

    // Pillar card — JSX frosted style: near-white fill, very subtle border
    slide.addShape("roundRect", {
      x: px, y: 2.95, w: pillarW, h: 3.65, rectRadius: 0.15,
      fill: { color: ac.fill, transparency: 93 },
      line: { color: ac.fill, width: 0.75 }
    });

    // Pillar title — filled pill
    addFilledPill(slide, pillar.title || "", px + 0.1, 3.08, pillarW - 0.2, 0.3, acc, 9);

    // Pillar body
    slide.addText(pillar.body || "", {
      x: px + 0.12, y: 3.5, w: pillarW - 0.24, h: 2.9,
      fontSize: 10, fontFace: FONT, color: C.dark,
      valign: "top", margin: 0
    });

    // Ghost sub-labels if provided
    if (pillar.tags) {
      pillar.tags.forEach((tag, ti) => {
        addGhostPill(slide, tag, px + 0.12, 5.7 + ti * 0.32, pillarW - 0.24, 0.24, acc, 8);
      });
    }
  });
}

// MATRIX GRID — page 14 of PDF
// Multi-row × multi-column grid with header row and label column
function addMatrixSlide(pres, data) {
  const slide = pres.addSlide();

  slide.addImage({ data: gradientStrip, x: 0, y: 6.92, w: W, h: 0.806 });

  addSectionLabel(slide, data.section || "");
  addXBlack(slide);

  addHeadline(slide, [
    { text: data.boldWord || "", bold: true },
    { text: (data.rest ? "\n" + data.rest : ""), bold: false }
  ], { x: ML, y: 0.48, w: 6.0, h: 1.3, size: 28 });

  const headers = data.headers || [];
  const rows    = data.rows    || [];
  const accs    = ["purple","blue","green","orange"];
  const labelColW = 1.5;
  const colW    = (W - ML - labelColW - 0.1) / Math.max(headers.length, 1);
  const rowH    = (H - 2.3) / Math.max(rows.length + 1, 1);
  const gridTopY = 2.25;

  // Header row — filled pills
  headers.forEach((h, i) => {
    const hx = ML + labelColW + i * colW;
    const pillH = 0.34;
    addFilledPill(slide, h, hx, gridTopY + (rowH - pillH) / 2, colW - 0.08, pillH, accs[i % accs.length], 10);
  });

  // Cap row height so content never overflows into bottom gradient strip (y=6.92)
  const maxGridBot = 6.88;
  const availH     = maxGridBot - gridTopY - rowH; // subtract header row
  const cappedRowH = Math.min(rowH, availH / Math.max(rows.length, 1));

  // Data rows
  rows.forEach((row, ri) => {
    const ry = gridTopY + cappedRowH * (ri + 1);
    const bg = ri % 2 === 0 ? "F8F8F8" : C.white;

    // Row background
    slide.addShape("rect", {
      x: ML, y: ry, w: W - ML * 2, h: cappedRowH,
      fill: { color: bg }, line: { color: "EEEEEE", width: 0.5 }
    });

    const lblPillH = 0.26;
    addGhostPill(slide, row.label || "", ML, ry + (cappedRowH - lblPillH) / 2, labelColW - 0.1, lblPillH, "cg", 8);

    (row.cells || []).forEach((cell, ci) => {
      const cx = ML + labelColW + ci * colW;
      // Centre-aligned to match header lozenges above
      slide.addText(cell, {
        x: cx + 0.08, y: ry + 0.06, w: colW - 0.16, h: cappedRowH - 0.12,
        fontSize: 10, fontFace: FONT, color: C.dark,
        align: "center", valign: "middle", margin: 0
      });
    });
  });
}



// IPAD VIDEO SLIDE
// Vector iPad in landscape, screen placeholder for video/image drop-in
// Left panel: section label + headline + caption. Right: iPad frame.
function addIpadVideoSlide(pres, data) {
  const slide = pres.addSlide();
  const splitX = 3.8;

  // Full-width white background base
  // gradientWide spans the full slide width behind the iPad — placed low so
  // the top stays white and the gradient sweeps in from the bottom
  slide.addImage({ data: gradientWide, x: 0, y: 3.75, w: W, h: H - 3.75 });

  // White overlay on left text panel only — keeps headline/caption readable
  slide.addShape("rect", {
    x: 0, y: 0, w: splitX, h: 3.75, fill: { color: C.white }, line: { color: C.white }
  });

  addSectionLabel(slide, data.section || "");
  addXBlack(slide);

  addHeadline(slide, [
    { text: data.boldWord || "", bold: true },
    { text: (data.rest ? "\n" + data.rest : ""), bold: false }
  ], { x: ML, y: 0.52, w: splitX - ML - 0.2, h: 2.2, size: 28 });

  if (data.caption) {
    slide.addText(data.caption, {
      x: ML, y: 6.1, w: splitX - ML - 0.2, h: 0.9,
      fontSize: 10, fontFace: FONT, color: C.dark,
      italic: true, valign: "bottom", margin: 0
    });
  }

  // iPad frame — landscape, fills right panel with comfortable margin
  const rightW = W - splitX;
  const padX   = 0.35;
  const padY   = 0.55;
  const dw     = rightW - padX * 2;
  const dh     = dw * (3 / 4);  // iPad landscape 4:3
  const dx     = splitX + padX;
  const dy     = (H - dh) / 2;

  const bezel  = dw * 0.03;
  const radius = dw * 0.035;
  const homeW  = dh * 0.032;  // home bar width
  const homeH  = dh * 0.008;

  // Outer frame — dark rounded rect
  slide.addShape("roundRect", {
    x: dx, y: dy, w: dw, h: dh, rectRadius: radius,
    fill: { color: "1C1C1E" }, line: { color: "3A3A3C", width: 1.5 }
  });

  // Screen area
  const sx = dx + bezel;
  const sy = dy + bezel;
  const sw = dw - bezel * 2;
  const sh = dh - bezel * 2;

  // Camera dot — top centre
  slide.addShape("ellipse", {
    x: dx + dw / 2 - 0.035, y: dy + bezel * 0.35, w: 0.07, h: 0.07,
    fill: { color: "2C2C2E" }, line: { color: "3A3A3C", width: 0.5 }
  });

  // Home bar — bottom centre
  slide.addShape("roundRect", {
    x: dx + (dw - homeW) / 2, y: dy + dh - bezel * 0.7, w: homeW, h: homeH,
    rectRadius: homeH / 2,
    fill: { color: "3A3A3C" }, line: { color: "3A3A3C", width: 0 }
  });

  // Side button — right edge
  slide.addShape("rect", {
    x: dx + dw, y: dy + dh * 0.32, w: 0.04, h: dh * 0.12,
    fill: { color: "3A3A3C" }, line: { color: "3A3A3C", width: 0 }
  });

  // Volume buttons — left edge
  slide.addShape("rect", {
    x: dx - 0.04, y: dy + dh * 0.25, w: 0.04, h: dh * 0.09,
    fill: { color: "3A3A3C" }, line: { color: "3A3A3C", width: 0 }
  });
  slide.addShape("rect", {
    x: dx - 0.04, y: dy + dh * 0.37, w: 0.04, h: dh * 0.09,
    fill: { color: "3A3A3C" }, line: { color: "3A3A3C", width: 0 }
  });

  // Screen placeholder — magenta video instruction text
  // USAGE RULE: ipad_video and phone_landscape are always used as a PAIR.
  // Both slides carry identical copy — user deletes whichever device shape they do not want.
  slide.addImage({ data: imgPlaceholder, x: sx, y: sy, w: sw, h: sh,
    sizing: { type: "cover", w: sw, h: sh } });

  // ONE text box — one click to delete
  const MV = "EC4899";
  slide.addText([
    { text: "Insert video here — DELETE AFTER USE\n", options: { bold: true,  fontSize: 8, fontFace: FONT, color: MV } },
    { text: "Use this slide for 4:3 video content.\n", options: { bold: false, fontSize: 8, fontFace: FONT, color: MV } },
    { text: "See following slide for 16:9 widescreen alternative.", options: { bold: false, fontSize: 8, fontFace: FONT, color: MV } }
  ], { x: sx + 0.1, y: sy + sh * 0.25, w: sw - 0.2, h: sh * 0.5,
       align: "center", valign: "top", margin: 0 });
}


// PHONE LANDSCAPE VIDEO SLIDE
// A single iPhone rendered in landscape orientation, maximising the 16:9 screen area.
// Left panel: section label + headline + caption (same structure as ipad_video).
// Right panel: gradientWide full width behind the device, phone centred in the space.
// The phone is drawn using the existing drawIPhoneFrame helper, rotated 90° via layout.
// Since pptxgenjs shapes can't rotate individually, we build landscape geometry directly:
//   - Outer rounded rect wider than tall (phone on its side)
//   - Home bar on the right edge, buttons on the top edge
//   - Notch/pill on the left edge (where the top speaker would be in landscape)
//   - Screen proportioned 16:9
function addPhoneLandscapeVideoSlide(pres, data) {
  const slide = pres.addSlide();
  const splitX = 3.8;

  // Full-width white base
  // gradientWide behind phone — full width, lower half
  slide.addImage({ data: gradientWide, x: 0, y: 3.75, w: W, h: H - 3.75 });

  // White overlay on left text area only
  slide.addShape("rect", {
    x: 0, y: 0, w: splitX, h: 3.75, fill: { color: C.white }, line: { color: C.white }
  });

  addSectionLabel(slide, data.section || "");
  addXBlack(slide);

  addHeadline(slide, [
    { text: data.boldWord || "", bold: true },
    { text: (data.rest ? "\n" + data.rest : ""), bold: false }
  ], { x: ML, y: 0.52, w: splitX - ML - 0.2, h: 2.2, size: 28 });

  if (data.caption) {
    slide.addText(data.caption, {
      x: ML, y: 6.1, w: splitX - ML - 0.2, h: 0.9,
      fontSize: 10, fontFace: FONT, color: C.dark,
      italic: true, valign: "bottom", margin: 0
    });
  }

  // Landscape phone geometry — centred in right panel with generous margin
  const rightX  = splitX + 0.3;
  const rightW  = W - splitX - 0.3;
  const dw      = rightW - 0.4;          // phone width (landscape = long axis)
  const dh      = dw * (9 / 19.5);      // phone height (landscape = short axis)
  const dx      = splitX + (W - splitX - dw) / 2;
  const dy      = (H - dh) / 2;

  const bezel   = dh * 0.07;            // bezel is fraction of short side
  const radius  = dh * 0.1;

  // Outer frame — landscape rounded rect
  slide.addShape("roundRect", {
    x: dx, y: dy, w: dw, h: dh, rectRadius: radius,
    fill: { color: "1A1A1A" }, line: { color: "444444", width: 1.5 }
  });

  // Screen area
  const sx = dx + bezel * 2.5;          // wider bezel on left (notch side)
  const sy = dy + bezel;
  const sw = dw - bezel * 3.5;          // narrower right (home bar side)
  const sh = dh - bezel * 2;
  slide.addShape("rect", {
    x: sx, y: sy, w: sw, h: sh,
    fill: { color: "111111" }, line: { color: "222222", width: 0.5 }
  });

  // Notch pill — left edge centre (speaker/camera in landscape)
  const notchH = dh * 0.35;
  const notchW = dh * 0.022;
  slide.addShape("roundRect", {
    x: dx + bezel * 0.6, y: dy + (dh - notchH) / 2,
    w: notchW, h: notchH, rectRadius: notchW / 2,
    fill: { color: "1A1A1A" }, line: { color: "1A1A1A", width: 0 }
  });

  // Home bar — right edge centre
  const homeH = dh * 0.3;
  const homeW = dh * 0.012;
  slide.addShape("roundRect", {
    x: dx + dw - bezel * 0.55 - homeW, y: dy + (dh - homeH) / 2,
    w: homeW, h: homeH, rectRadius: homeW / 2,
    fill: { color: "3A3A3C" }, line: { color: "3A3A3C", width: 0 }
  });

  // Volume buttons — top edge
  slide.addShape("rect", {
    x: dx + dw * 0.28, y: dy - 0.04, w: dw * 0.07, h: 0.04,
    fill: { color: "333333" }, line: { color: "333333", width: 0 }
  });
  slide.addShape("rect", {
    x: dx + dw * 0.38, y: dy - 0.04, w: dw * 0.07, h: 0.04,
    fill: { color: "333333" }, line: { color: "333333", width: 0 }
  });

  // Power button — bottom edge
  slide.addShape("rect", {
    x: dx + dw * 0.32, y: dy + dh, w: dw * 0.08, h: 0.04,
    fill: { color: "333333" }, line: { color: "333333", width: 0 }
  });

  // Screen placeholder — magenta video instruction text
  // USAGE RULE: phone_landscape and ipad_video are always used as a PAIR.
  // Both slides carry identical copy — user deletes whichever device shape they do not want.
  slide.addImage({ data: imgPlaceholder, x: sx, y: sy, w: sw, h: sh,
    sizing: { type: "cover", w: sw, h: sh } });

  // ONE text box — one click to delete
  const MPV = "EC4899";
  slide.addText([
    { text: "Insert video here — DELETE AFTER USE\n",      options: { bold: true,  fontSize: 8, fontFace: FONT, color: MPV } },
    { text: "Use this slide for 16:9 widescreen video content.\n", options: { bold: false, fontSize: 8, fontFace: FONT, color: MPV } },
    { text: "See previous slide for 4:3 alternative.",      options: { bold: false, fontSize: 8, fontFace: FONT, color: MPV } }
  ], { x: sx + 0.1, y: sy + sh * 0.15, w: sw - 0.2, h: sh * 0.6,
       align: "center", valign: "top", margin: 0 });
}

// FULL BLEED CAPTION SLIDE
// Full-bleed image placeholder with white caption box bottom-left.
// Caption box height is DYNAMIC — expands to fit the supplied text.
// X icon always white — sits over full-bleed image.
// Magenta ChatGPT + Google prompts on image placeholder.
function addFullBleedCaptionSlide(pres, data) {
  const slide = pres.addSlide();

  const ratio  = data.imageRatio  || "16:9";
  const basePromptFBC = data.imagePrompt || "A dramatic, high-quality brand experience or live event photograph. Dark, moody atmosphere with vibrant accent lighting. Premium, editorial feel. No text overlays.";
  const prompt = basePromptFBC.includes(ratio) ? basePromptFBC : basePromptFBC.replace(/[.\s]+$/, "") + ". Generate this image in 16:9 widescreen landscape format.";
  const google = data.googleSearch || "site:unsplash.com brand activation live event dramatic lighting";

  // Dark underlay
  slide.addShape("rect", {
    x: 0, y: 0, w: W, h: H, fill: { color: "1A1A1A" }, line: { color: "1A1A1A" }
  });

  // Image placeholder with magenta prompts
  slide.addImage({
    data: imgPlaceholder,
    x: 0, y: 0, w: W, h: H,
    sizing: { type: "cover", w: W, h: H }
  });

  // Magenta prompts — one text box, one click to delete
  // Google search text is a live hyperlink
  const MFC  = "EC4899";
  const gUrlFBC = googleImagesUrl(google);
  slide.addText([
    { text: "ChatGPT image prompt — DELETE AFTER USE:\n", options: { bold: true,  fontSize: 10, fontFace: FONT, color: MFC } },
    { text: prompt + "\n\n",                             options: { bold: false, fontSize: 10, fontFace: FONT, color: MFC } },
    { text: "Google Images search — click to open, DELETE AFTER USE:\n", options: { bold: true, fontSize: 10, fontFace: FONT, color: MFC } },
    { text: google, options: { bold: false, fontSize: 10, fontFace: FONT, color: MFC, hyperlink: { url: gUrlFBC } } }
  ], { x: 5.5, y: 1.4, w: 7.4, h: 2.2, align: "center", valign: "top", margin: 0 });

  // X icon always white over image
  addXWhite(slide);

  // Section label top-left — white
  addSectionLabel(slide, data.section || "", { white: true });

  // White caption box — DYNAMIC height based on content
  // Minimum height fits headline; expands if subcaption is also present
  const capW   = data.captionWidth || 5.5;
  const capX   = 0;
  const lineHt  = 0.52;   // approx height per line at 22pt
  const hlLines = data.headline ? data.headline.split("\n").length : 0;
  const hasSubcap = !!data.subcaption;
  // Dynamic: top padding 0.22" + headline lines + gap + subcaption if present + bottom padding 0.22"
  const capH   = 0.22 + (hlLines * lineHt) + (hasSubcap ? 0.18 + 0.38 : 0) + 0.22;
  const capY   = H - capH;

  slide.addShape("rect", {
    x: capX, y: capY, w: capW, h: capH,
    fill: { color: C.white }, line: { color: C.white }
  });

  // Headline — line 1 bold, line 2 regular. Padding 0.22" from top of box.
  if (data.headline) {
    const lines = data.headline.split("\n");
    const runs  = [];
    lines.forEach((line, i) => {
      if (i > 0) runs.push({ text: "\n", options: { fontSize: 22, fontFace: FONT, color: C.dark } });
      runs.push({ text: line, options: { bold: i === 0, fontSize: 22, fontFace: FONT, color: C.dark } });
    });
    slide.addText(runs, {
      x: capX + ML, y: capY + 0.22, w: capW - ML - 0.2, h: hlLines * lineHt,
      valign: "top", margin: 0
    });
  }

  // Subcaption — 10pt, grey, below headline with 0.18" gap
  if (data.subcaption) {
    slide.addText(data.subcaption, {
      x: capX + ML, y: capY + 0.22 + hlLines * lineHt + 0.18, w: capW - ML - 0.2, h: 0.38,
      fontSize: 10, fontFace: FONT, color: C.midGrey,
      valign: "top", margin: 0
    });
  }
}

// APPENDIX DIVIDER — alias of standard divider
// Use type: "appendix_divider" to mark the start of supporting/reference material
function addAppendixDividerSlide(pres, data) {
  addDividerSlide(pres, { ...data, title: data.title || "Appendix" });
}

// ============================================================
// CHART SLIDE FUNCTIONS
// Styled using the CG JSX design system (CG_Sales_Intelligence_v40.jsx):
//   - Colours: 9333ea (purple), ec4899 (pink), 10b981 (green), f97316 (orange)
//   - Frosted glass insight pill: white 82% opacity, blur(16px), coloured border
//   - Section labels: 9px, 0.16em tracking, 700 weight, uppercase, colour-matched
//   - Headline: 800 weight, -0.02em tracking (JSX heading style)
//   - Grid lines: E8E8E8 (JSX chart grid)
//   - Insight lozenge: pill shape (borderRadius 100), colour+15 tint bg, colour+40 border
// ============================================================

// JSX colour palette — pulled directly from MODULES array in the JSX
const CHART_COLORS = [
  "C4943A",  // CME gold — primary series
  "ec4899",  // outreach pink — secondary series
  "10b981",  // marcomms green — tertiary
  "f97316",  // sense orange — quaternary
  "A87D2E",  // CME dark gold — fifth
  "db2777",  // deeper pink — sixth
];

// Shared chart options — JSX grid style
function chartBaseOpts(extra = {}) {
  return {
    catAxisLabelFontSize: 10,
    catAxisLabelFontFace: FONT,
    valAxisLabelFontSize: 10,
    valAxisLabelFontFace: FONT,
    legendFontSize: 10,
    legendFontFace: FONT,
    catGridLine: { style: "none" },
    valGridLine: { color: "E8E8E8", style: "solid" },
    ...extra
  };
}

// JSX-styled chart slide base
// - White background + gradient strip flush to bottom
// - Section label: 9px uppercase micro-label, left-aligned with ML
// - Headline uses JSX heading style: fontWeight 800, tight tracking
// - Insight lozenge: pill above gradient, not overlapping it
// - Chart occupies LEFT TWO-THIRDS; right third reserved for commentary text
function chartSlideBase(pres, data) {
  const slide = pres.addSlide();

  // White background
  // Solid colour underlay behind gradient strip — catches any sub-pixel gap
  slide.addShape(pres.ShapeType.rect, {
    x: 0, y: 6.92, w: W, h: 0.806,
    fill: { color: "C7D2FE" }, line: { color: "C7D2FE" }
  });
  // Gradient strip — white border cropped at source, exact cover flush to bottom
  slide.addImage({ data: gradientStrip, x: 0, y: 6.92, w: W, h: 0.806 });

  // Section label — left-aligned with master margin, JSX micro-label style
  if (data.section) {
    slide.addText(data.section.toUpperCase(), {
      x: ML, y: 0.16, w: 4.0, h: 0.22,
      fontSize: 9, fontFace: FONT, color: CHART_COLORS[0],
      bold: true, charSpacing: 2.2
    });
  }

  addXBlack(slide);

  // Headline — JSX heading style: 800 weight, tight tracking
  if (data.boldWord || data.title) {
    const headRuns = data.boldWord
      ? [
          { text: data.boldWord, options: { bold: true, fontSize: 28, fontFace: FONT, color: C.dark, charSpacing: -0.5 } },
          { text: data.rest ? ("\n" + data.rest) : "", options: { bold: false, fontSize: 28, fontFace: FONT, color: "888888", charSpacing: -0.5 } }
        ]
      : [{ text: data.title, options: { bold: true, fontSize: 28, fontFace: FONT, color: C.dark, charSpacing: -0.5 } }];
    slide.addText(headRuns, {
      x: ML, y: 0.44, w: 10.5, h: 1.3, valign: "top", margin: 0
    });
  }

  // Insight pill — sits above gradient strip, clear of it
  if (data.insight) {
    slide.addShape(pres.ShapeType.roundRect, {
      x: ML, y: 6.44, w: W - ML - 0.3, h: 0.36,
      fill: { color: "F3E8FF" },
      line: { color: "D8B4FE", width: 0.75 },
      rectRadius: 0.18
    });
    slide.addText(data.insight, {
      x: ML + 0.16, y: 6.44, w: W - ML - 0.62, h: 0.36,
      fontSize: 10, fontFace: FONT, color: "6B21A8",
      italic: false, bold: false, valign: "middle", margin: 0
    });
  }

  return slide;
}

// Layout constants for chart slides
// Charts sit in the left 2/3 of the slide; right 1/3 is commentary
const CHART_W   = 8.4;   // left two-thirds width (of 13.33" total)
const CHART_X   = ML;    // starts at master left margin
const COMMENT_X = 9.0;   // commentary column starts here
const COMMENT_W = 4.0;   // commentary column width
const COMMENT_Y = 2.0;   // aligns with chart top
const COMMENT_H = 4.3;   // matches chart height

// Helper: add commentary column — AI MUST always populate this with real insight copy.
// "Add commentary here." must NEVER appear in a generated deck.
// If data.commentary is missing, generator logs a warning — do not leave it empty.
// RULE: Every chart slide requires a commentary or insight field in the BRIEF.
function addCommentaryColumn(pres, slide, text) {
  // Subtle left border line
  slide.addShape(pres.ShapeType.line, {
    x: COMMENT_X - 0.18, y: COMMENT_Y, w: 0, h: COMMENT_H,
    line: { color: "E8E8E8", width: 1.0 }
  });
  if (!text) {
    console.warn("CHART SLIDE: commentary field is empty — AI must always supply insight copy. Add a commentary or insight field to the BRIEF.");
    return;
  }
  slide.addText(text, {
    x: COMMENT_X, y: COMMENT_Y, w: COMMENT_W, h: COMMENT_H,
    fontSize: 11.5, fontFace: FONT, color: C.dark,
    wrap: true, valign: "top", margin: 0
  });
}

// --- BAR CHART (vertical columns) ---
// data.labels: ["A","B","C"]
// data.series: [{ name, values:[] }]
// data.showValue: true to print values on bars
// data.commentary: optional text for right-hand column
function addBarChartSlide(pres, data) {
  const slide = chartSlideBase(pres, data);
  const series = (data.series || []).map((s, i) => ({
    name: s.name, labels: s.labels || data.labels || [], values: s.values || [],
    color: s.color || CHART_COLORS[i % CHART_COLORS.length]
  }));
  slide.addChart(pres.ChartType.bar, series, chartBaseOpts({
    x: CHART_X, y: COMMENT_Y, w: CHART_W, h: COMMENT_H,
    barDir: "col",
    barGapWidthPct: 60,
    chartColors: series.map(s => s.color),
    showLegend: series.length > 1,
    legendPos: "b",
    showValue: !!data.showValue,
    dataLabelFontSize: 9,
    dataLabelFontFace: FONT,
  }));
  addCommentaryColumn(pres, slide, data.commentary);
}

// --- HORIZONTAL BAR CHART ---
// data.labels: ["Category A","Category B"]
// data.series: [{ name, values:[] }]
// data.commentary: optional text for right-hand column
function addHorizBarChartSlide(pres, data) {
  const slide = chartSlideBase(pres, data);
  const series = (data.series || []).map((s, i) => ({
    name: s.name, labels: s.labels || data.labels || [], values: s.values || [],
    color: s.color || CHART_COLORS[i % CHART_COLORS.length]
  }));
  slide.addChart(pres.ChartType.bar, series, chartBaseOpts({
    x: CHART_X, y: COMMENT_Y, w: CHART_W, h: COMMENT_H,
    barDir: "bar",
    barGapWidthPct: 40,
    chartColors: series.map(s => s.color),
    showLegend: series.length > 1,
    legendPos: "b",
    showValue: !!data.showValue,
    dataLabelFontSize: 9,
    dataLabelFontFace: FONT,
  }));
  addCommentaryColumn(pres, slide, data.commentary);
}

// --- LINE CHART ---
// data.labels: ["Q1","Q2","Q3"]
// data.series: [{ name, values:[] }]
function addLineChartSlide(pres, data) {
  const slide = chartSlideBase(pres, data);
  const series = (data.series || []).map((s, i) => ({
    name: s.name, labels: s.labels || data.labels || [], values: s.values || [],
    color: s.color || CHART_COLORS[i % CHART_COLORS.length]
  }));
  slide.addChart(pres.ChartType.line, series, chartBaseOpts({
    x: CHART_X, y: COMMENT_Y, w: CHART_W, h: COMMENT_H,
    chartColors: series.map(s => s.color),
    lineDataSymbol: "none",
    lineSize: 3,
    showLegend: series.length > 1,
    legendPos: "b",
  }));
  addCommentaryColumn(pres, slide, data.commentary);
}

// --- PIE CHART ---
// data.chartData: [{ label, value }]
// data.commentary: optional text for right-hand column
function addPieChartSlide(pres, data) {
  const slide = chartSlideBase(pres, data);
  const items = data.chartData || [];
  const chartData = [{ name: data.title || "Data", labels: items.map(d => d.label), values: items.map(d => d.value) }];
  const colors = items.map((d, i) => d.color || CHART_COLORS[i % CHART_COLORS.length]);
  // Pie: slightly smaller than full chart area to keep proportions good
  slide.addChart(pres.ChartType.pie, chartData, {
    x: CHART_X, y: COMMENT_Y, w: CHART_W, h: COMMENT_H,
    chartColors: colors,
    showLegend: true,
    legendPos: "b",
    legendFontSize: 10,
    legendFontFace: FONT,
    showLabel: true,
    showPercent: true,
    dataLabelFontSize: 10,
    dataLabelFontFace: FONT,
    dataLabelColor: C.white,
  });
  addCommentaryColumn(pres, slide, data.commentary || data.insight);
}

// --- DOUGHNUT CHART ---
// data.chartData: [{ label, value }]
// data.commentary: optional text for right-hand column
function addDoughnutChartSlide(pres, data) {
  const slide = chartSlideBase(pres, data);
  const items = data.chartData || [];
  const chartData = [{ name: data.title || "Data", labels: items.map(d => d.label), values: items.map(d => d.value) }];
  const colors = items.map((d, i) => d.color || CHART_COLORS[i % CHART_COLORS.length]);
  slide.addChart(pres.ChartType.doughnut, chartData, {
    x: CHART_X, y: COMMENT_Y, w: CHART_W, h: COMMENT_H,
    chartColors: colors,
    holeSize: 55,
    showLegend: true,
    legendPos: "b",
    legendFontSize: 10,
    legendFontFace: FONT,
    showLabel: true,
    showPercent: true,
    dataLabelFontSize: 10,
    dataLabelFontFace: FONT,
    dataLabelColor: C.white,
  });
  addCommentaryColumn(pres, slide, data.commentary || data.insight);
}

// --- COMBO CHART (bars + line overlay) ---
// data.labels: ["Jan","Feb"]
// data.barSeries:  [{ name, values }]
// data.lineSeries: [{ name, values }]
// data.commentary: optional text for right-hand column
function addComboChartSlide(pres, data) {
  const slide = chartSlideBase(pres, data);
  const barSeries = (data.barSeries || []).map((s, i) => ({
    name: s.name, labels: s.labels || data.labels || [], values: s.values || [],
    color: s.color || CHART_COLORS[i]
  }));
  const lineSeries = (data.lineSeries || []).map((s, i) => ({
    name: s.name, labels: s.labels || data.labels || [], values: s.values || [],
    color: s.color || CHART_COLORS[2 + i]
  }));
  slide.addChart(
    [
      { type: pres.ChartType.bar,  data: barSeries,
        options: { chartColors: barSeries.map(s => s.color), barDir: "col", barGapWidthPct: 60 } },
      { type: pres.ChartType.line, data: lineSeries,
        options: { chartColors: lineSeries.map(s => s.color), lineDataSymbol: "none", lineSize: 3, secondaryValAxis: false } }
    ],
    chartBaseOpts({
      x: CHART_X, y: COMMENT_Y, w: CHART_W, h: COMMENT_H,
      showLegend: true,
      legendPos: "b",
    })
  );
  addCommentaryColumn(pres, slide, data.commentary);
}

// ============================================================
// TWO-COLUMN TEXT SLIDE
// White bg, headline top-left, body split into two equal columns
// 1cm (~0.39") central gutter matching CG PowerPoint padding standard
// Each column can have an optional accent colour pill label
// ============================================================
function addTwoColumnTextSlide(pres, data) {
  const slide = pres.addSlide();

  slide.addImage({ data: gradientStrip, x: 0, y: 6.92, w: W, h: 0.806 });

  addSectionLabel(slide, data.section || "");
  addXBlack(slide);

  // Headline
  addHeadline(slide, [
    { text: data.boldWord || "", bold: true },
    { text: data.rest ? "\n" + data.rest : "", bold: false }
  ], { x: ML, y: 0.48, w: 10.0, h: 1.3, size: 32 });

  // Two columns — 1cm (0.39") central gutter
  const gutter  = 0.39;
  const colW    = (W - ML * 2 - gutter) / 2;
  const col1X   = ML;
  const col2X   = ML + colW + gutter;
  const bodyY   = 2.1;
  const bodyH   = 4.5;

  const cols = data.columns || [];

  // 15px (~0.21") extra breathing space added between headline and pill/rule/body
  const contentY = 2.31;  // was 2.1 — shifted down by ~0.21"

  cols.slice(0, 2).forEach((col, i) => {
    const x = i === 0 ? col1X : col2X;

    // Accent label pill above rule
    if (col.label) {
      const acc = col.accent || "purple";
      addFilledPill(slide, col.label, x, contentY - 0.38, colW * 0.55, 0.28, acc, 9);
    }

    // Accent rule
    const ac = ACCENT[col.accent || "cg"];
    slide.addShape(pres.ShapeType.line, {
      x, y: contentY, w: colW, h: 0,
      line: { color: ac.fill, width: 1.5 }
    });

    // Body text — starts 0.18" below rule
    const runs = parseInlineBold(col.body || "", 12, C.dark);
    slide.addText(runs, {
      x, y: contentY + 0.18, w: colW, h: H - contentY - 0.18 - 0.9,
      fontSize: 12, fontFace: FONT, color: C.dark,
      valign: "top", margin: 0
    });
  });
}

// ============================================================
// THREE-COLUMN TEXT SLIDE (new style)
// White bg, headline, three equal columns with 1cm gutters
// Each column: accent pill label + rule + body copy
// ============================================================
function addThreeColumnTextSlide(pres, data) {
  const slide = pres.addSlide();

  slide.addImage({ data: gradientStrip, x: 0, y: 6.92, w: W, h: 0.806 });

  addSectionLabel(slide, data.section || "");
  addXBlack(slide);

  addHeadline(slide, [
    { text: data.boldWord || "", bold: true },
    { text: data.rest ? "\n" + data.rest : "", bold: false }
  ], { x: ML, y: 0.48, w: 10.0, h: 1.3, size: 32 });

  const gutter = 0.39;  // 1cm
  const colW   = (W - ML * 2 - gutter * 2) / 3;
  const bodyY  = 2.1;
  const bodyH  = 4.5;
  const accentKeys = ["purple", "green", "orange"];

  // 15px (~0.21") extra breathing space between headline and pill/rule/body
  const content3Y = 2.31;  // was 2.1

  (data.columns || []).slice(0, 3).forEach((col, i) => {
    const x   = ML + i * (colW + gutter);
    const acc = col.accent || accentKeys[i] || "cg";
    const ac  = ACCENT[acc];

    // Accent pill label
    if (col.label) {
      addFilledPill(slide, col.label, x, content3Y - 0.38, colW * 0.72, 0.28, acc, 9);
    }

    // Accent rule
    slide.addShape(pres.ShapeType.line, {
      x, y: content3Y, w: colW, h: 0,
      line: { color: ac.fill, width: 1.5 }
    });

    // Body text
    const runs = parseInlineBold(col.body || "", 11.5, C.dark);
    slide.addText(runs, {
      x, y: content3Y + 0.18, w: colW, h: H - content3Y - 0.18 - 0.9,
      fontSize: 11.5, fontFace: FONT, color: C.dark,
      valign: "top", margin: 0
    });
  });
}

// ============================================================
// DATA TABLE SLIDE
// Renders a data table (e.g. from Excel) using the CG colour system.
// Header row: filled accent pills. Alternating row tints.
// Optional section grouping rows with ghost pills.
// Right panel: key stats pulled from the data.
// ============================================================
function addDataTableSlide(pres, data) {
  const slide = pres.addSlide();

  slide.addImage({ data: gradientStrip, x: 0, y: 6.92, w: W, h: 0.806 });

  addSectionLabel(slide, data.section || "");
  addXBlack(slide);

  addHeadline(slide, [
    { text: data.boldWord || "", bold: true },
    { text: data.rest ? "\n" + data.rest : "", bold: false }
  ], { x: ML, y: 0.48, w: 8.5, h: 1.3, size: 28 });

  const rows      = data.rows    || [];
  const headers   = data.headers || [];
  const tableAccent = data.accent || "purple";
  const ac        = ACCENT[tableAccent];

  // Table layout
  const tableX    = ML;
  const tableW    = data.statsPanel ? 8.6 : W - ML * 2;
  const startY    = 2.1;
  const headerH   = 0.38;
  const rowH      = data.rowHeight || ((H - startY - 0.9 - headerH) / Math.max(rows.length, 1));
  const cappedRowH = Math.min(rowH, 0.52);

  // Column widths — first col label wider
  const labelColW = data.labelColW || 2.8;
  const dataColW  = (tableW - labelColW) / Math.max(headers.length, 1);

  // Header row — filled pills for each column header
  const headerAccents = ["purple", "blue", "green", "orange", "pink", "cg"];
  headers.forEach((h, i) => {
    const hx = tableX + labelColW + i * dataColW;
    addFilledPill(slide, h, hx + 0.05, startY + (headerH - 0.3) / 2, dataColW - 0.1, 0.3, headerAccents[i % headerAccents.length], 9);
  });

  // Empty label col header — ghost pill
  if (data.labelHeader) {
    addGhostPill(slide, data.labelHeader, tableX, startY + (headerH - 0.3) / 2, labelColW - 0.1, 0.3, tableAccent, 9);
  }

  // Data rows
  rows.forEach((row, ri) => {
    const ry  = startY + headerH + ri * cappedRowH;
    const bg  = ri % 2 === 0 ? ac.tint : C.white;

    // Row background
    slide.addShape(pres.ShapeType.rect, {
      x: tableX, y: ry, w: tableW, h: cappedRowH,
      fill: { color: bg },
      line: { color: "EEEEEE", width: 0.5 }
    });

    // Row label — ghost pill if it's a section header, plain text otherwise
    if (row.isHeader) {
      addFilledPill(slide, row.label || "", tableX, ry + (cappedRowH - 0.24) / 2, labelColW - 0.1, 0.24, tableAccent, 8);
    } else {
      slide.addText(row.label || "", {
        x: tableX + 0.1, y: ry + 0.04, w: labelColW - 0.15, h: cappedRowH - 0.08,
        fontSize: 10, fontFace: FONT, color: C.dark,
        valign: "middle", margin: 0
      });
    }

    // Data cells
    (row.cells || []).forEach((cell, ci) => {
      const cx = tableX + labelColW + ci * dataColW;
      const isBold = row.isBold || false;
      slide.addText(String(cell), {
        x: cx + 0.08, y: ry + 0.04, w: dataColW - 0.12, h: cappedRowH - 0.08,
        fontSize: 10, fontFace: FONT, color: C.dark,
        bold: isBold, align: "right", valign: "middle", margin: 0
      });
    });
  });

  // Bottom border
  const tableBot = startY + headerH + rows.length * cappedRowH;
  slide.addShape(pres.ShapeType.line, {
    x: tableX, y: tableBot, w: tableW, h: 0,
    line: { color: ac.fill, width: 1.5 }
  });

  // Optional right stats panel
  if (data.statsPanel) {
    const panelX = 9.2;
    const panelW = W - panelX - ML;
    let sy = startY;

    (data.statsPanel || []).forEach((stat, i) => {
      const acc  = ["purple", "green", "orange", "blue", "pink"][i % 5];
      const sacc = ACCENT[acc];

      // Frosted card
      slide.addShape(pres.ShapeType.roundRect, {
        x: panelX, y: sy, w: panelW, h: 0.9,
        rectRadius: 0.08,
        fill: { color: sacc.fill, transparency: 90 },
        line: { color: sacc.fill, width: 0.75 }
      });

      // Big number
      slide.addText(stat.value, {
        x: panelX + 0.1, y: sy + 0.04, w: panelW - 0.2, h: 0.46,
        fontSize: 22, fontFace: FONT, bold: true, color: sacc.fill,
        align: "center", valign: "middle", margin: 0
      });

      // Label
      slide.addText(stat.label, {
        x: panelX + 0.1, y: sy + 0.5, w: panelW - 0.2, h: 0.32,
        fontSize: 8.5, fontFace: FONT, color: C.dark,
        align: "center", valign: "top", margin: 0
      });

      sy += 1.02;
    });
  }
}

// --- STATEMENT OF INTENT SLIDE --- //
// MANDATORY: this slide must appear as slide 2 in every deck — after title, before contents.
// Full-bleed image with bold white headline bottom-left. Maximum 3 lines.
// Always optimistic, forward-looking, punchy. Sets the tone for the entire document.
// AI must write a dynamic, specific statement derived from the deck purpose — never generic.
// Image should be bright, aspirational, high-energy — colourful and premium.
// X icon always white. All text bold and white.
function addStatementOfIntentSlide(pres, data) {
  const slide = pres.addSlide();

  const prompt = data.imagePrompt || "Bright, optimistic, high-energy brand experience photograph. Vibrant colours, premium production, aspirational energy. Wide angle, cinematic. No text overlays. Generate in 16:9 widescreen landscape format.";
  const google = data.googleSearch || "site:unsplash.com vibrant brand experience energetic crowd colourful";

  // Dark underlay so white text is always readable
  slide.addShape("rect", {
    x: 0, y: 0, w: W, h: H, fill: { color: "111111" }, line: { color: "111111" }
  });

  slide.addImage({ data: imgPlaceholder, x: 0, y: 0, w: W, h: H,
    sizing: { type: "cover", w: W, h: H } });

  // Magenta prompts — one text box, one click to delete
  // Google search text is a live hyperlink
  const MF  = "EC4899";
  const gUrlSOI = googleImagesUrl(google);
  slide.addText([
    { text: "ChatGPT image prompt — DELETE AFTER USE:\n", options: { bold: true,  fontSize: 10, fontFace: FONT, color: MF } },
    { text: prompt + "\n\n",                             options: { bold: false, fontSize: 10, fontFace: FONT, color: MF } },
    { text: "Google Images search — click to open, DELETE AFTER USE:\n", options: { bold: true, fontSize: 10, fontFace: FONT, color: MF } },
    { text: google, options: { bold: false, fontSize: 10, fontFace: FONT, color: MF, hyperlink: { url: gUrlSOI } } }
  ], { x: 6.0, y: 1.4, w: 7.0, h: 2.2, align: "center", valign: "top", margin: 0 });

  // X icon always white over image
  addXWhite(slide);

  // Statement headline — bold white, bottom-left, max 3 lines
  if (data.headline) {
    const lines = data.headline.split("\n").slice(0, 3);
    const runs  = [];
    lines.forEach((line, i) => {
      if (i > 0) runs.push({ text: "\n", options: { fontSize: 44, fontFace: FONT, color: C.white, bold: true } });
      runs.push({ text: line, options: { bold: true, fontSize: 44, fontFace: FONT, color: C.white } });
    });
    slide.addText(runs, {
      x: ML, y: 3.5, w: 11.0, h: 3.7,
      valign: "bottom", margin: 0
    });
  }
}


// --- CUSTOM DIAGRAM SLIDE --- //
// Use when none of the existing diagram types fits the story to be told.
// The AI writing the BRIEF describes the diagram using a draw array of primitive
// instructions executed at build time using the full CG design system.
//
// AVAILABLE PRIMITIVES (specify as fn + args in the draw array):
//   "filledPill"  — addFilledPill(slide, text, x, y, w, h, accentKey, fontSize)
//   "ghostPill"   — addGhostPill(slide, text, x, y, w, h, accentKey, fontSize)
//   "frostedBox"  — roundRect with accent fill at 90% transparency (x, y, w, h, accentKey, radius)
//   "roundRect"   — roundRect with accent fill (x, y, w, h, accentKey, transparency, radius)
//   "ellipse"     — ellipse with accent fill (x, y, w, h, accentKey, transparency)
//   "rect"        — rect with accent fill (x, y, w, h, accentKey, transparency)
//   "line"        — straight line (x, y, w, h, colorKey, width, dashStyle)
//   "arrow"       — line with arrowhead (x, y, w, h, colorKey, width)
//   "text"        — body copy (str, x, y, w, h, opts:{fontSize,bold,color,align,valign})
//   "watermarkX"  — ghosted X watermark (accentKey)
//
// ACCENT KEYS: "purple", "green", "pink", "orange", "blue", "cg"
// COLOR KEYS:  "dark", "white", "grey" (resolved from C object), or raw hex
// BACKGROUND:  "white" (default), "gradient", "split"
//
// AI DESIGN RULES:
//   - filledPill = primary labels; ghostPill = secondary/supporting
//   - Frosted box = roundRect accentKey, transparency 90, radius 0.06-0.12
//   - Connector lines: colorKey "CCCCCC" or "DDDDDD", width 0.75-1.5
//   - All coordinates in inches: W=13.33, H=7.5, ML=0.38
//   - Negative line widths are auto-corrected (no cx errors)
function addCustomDiagramSlide(pres, data) {
  const slide = pres.addSlide();
  const bg    = data.background || "white";

  // Background
  if (bg === "gradient") {
    slide.addShape("rect", { x: 0, y: 0, w: W, h: H,
      fill: { color: "BFC8F5" }, line: { color: "BFC8F5" } });
    slide.addImage({ data: gradientFull, x: 0, y: -0.123, w: W, h: 7.746 });

  } else if (bg === "split") {
    slide.addShape("rect", { x: 0, y: 0, w: W / 3, h: H,
      fill: { color: "FDF4FF" }, line: { color: "FDF4FF" } });
    slide.addImage({ data: gradientFull, x: W / 3, y: 0, w: W - W / 3, h: H });

  } else {
    // white — gradient strip at bottom
    slide.addImage({ data: gradientStrip, x: 0, y: 6.92, w: W, h: 0.806 });
  }

  const isLight = bg === "white";
  addSectionLabel(slide, data.section || "", { white: !isLight });
  isLight ? addXBlack(slide) : addXWhite(slide);

  if (data.boldWord || data.rest) {
    addHeadline(slide, [
      { text: data.boldWord || "", bold: true },
      { text: data.rest ? "\n" + data.rest : "", bold: false }
    ], {
      x: ML, y: 0.48, w: bg === "split" ? W / 3 - ML - 0.2 : W * 0.55, h: 1.3,
      size: 28, color: isLight ? C.dark : C.dark
    });
  }

  // Colour resolver — accepts accentKey or hex string
  function resolveColor(key) {
    if (!key) return C.dark;
    if (key === "dark")  return C.dark;
    if (key === "white") return C.white;
    if (key === "grey")  return C.grey;
    if (ACCENT[key])     return ACCENT[key].fill;
    return key;  // treat as raw hex
  }

  // Execute draw primitives
  const draws = data.draw || [];
  draws.forEach(d => {
    const a = d.args || [];
    switch (d.fn) {

      case "filledPill":
        // (text, x, y, w, h, accentKey, fontSize)
        addFilledPill(slide, a[0]||"", a[1]||0, a[2]||0, a[3]||1, a[4]||0.3, a[5]||"cg", a[6]||10);
        break;

      case "ghostPill":
        // (text, x, y, w, h, accentKey, fontSize)
        addGhostPill(slide, a[0]||"", a[1]||0, a[2]||0, a[3]||1, a[4]||0.3, a[5]||"cg", a[6]||10);
        break;

      case "frostedBox": {
        // (x, y, w, h, accentKey, radius)
        const ac = ACCENT[a[4]] || ACCENT.cg;
        slide.addShape("roundRect", {
          x: a[0]||0, y: a[1]||0, w: a[2]||1, h: a[3]||1,
          rectRadius: a[5] !== undefined ? a[5] : 0.12,
          fill: { color: ac.fill, transparency: 90 },
          line: { color: ac.fill, width: 0.75 }
        });
        break;
      }

      case "line": {
        // (x, y, w, h, color, width, dash)
        // Auto-correct negative w/h: flip start coordinate so cx/cy are always positive
        const dashType = a[6] === "dot" ? "sysDot" : a[6] === "dash" ? "dash" : "solid";
        let lx = a[0]||0, ly = a[1]||0, lw = a[2]||0, lh = a[3]||0;
        if (lw < 0) { lx = lx + lw; lw = -lw; }
        if (lh < 0) { ly = ly + lh; lh = -lh; }
        slide.addShape("line", {
          x: lx, y: ly, w: lw, h: lh,
          line: { color: resolveColor(a[4]) || "AAAAAA", width: a[5]||1,
                  dashType: dashType === "solid" ? undefined : dashType }
        });
        break;
      }

      case "arrow": {
        // (x, y, w, h, color, width)
        slide.addShape("line", {
          x: a[0]||0, y: a[1]||0, w: a[2]||1, h: a[3]||0,
          line: {
            color: resolveColor(a[4]) || "AAAAAA", width: a[5]||1.5,
            endArrowType: "triangle"
          }
        });
        break;
      }

      case "ellipse": {
        // (x, y, w, h, accentKey, transparency)
        const ac = ACCENT[a[4]] || ACCENT.cg;
        const trans = a[5] !== undefined ? a[5] : 80;
        slide.addShape("ellipse", {
          x: a[0]||0, y: a[1]||0, w: a[2]||1, h: a[3]||1,
          fill: { color: ac.fill, transparency: trans },
          line: { color: ac.fill, width: 0.75 }
        });
        break;
      }

      case "rect": {
        // (x, y, w, h, accentKey, transparency)
        const ac = ACCENT[a[4]] || ACCENT.cg;
        const trans = a[5] !== undefined ? a[5] : 80;
        slide.addShape("rect", {
          x: a[0]||0, y: a[1]||0, w: a[2]||1, h: a[3]||1,
          fill: { color: ac.fill, transparency: trans },
          line: { color: ac.fill, width: 0.75 }
        });
        break;
      }

      case "roundRect": {
        // (x, y, w, h, accentKey, transparency, radius)
        const ac = ACCENT[a[4]] || ACCENT.cg;
        const trans  = a[5] !== undefined ? a[5] : 90;
        const radius = a[6] !== undefined ? a[6] : 0.08;
        slide.addShape("roundRect", {
          x: a[0]||0, y: a[1]||0, w: a[2]||1, h: a[3]||1,
          rectRadius: radius,
          fill: { color: ac.fill, transparency: trans },
          line: { color: ac.fill, width: 0.75 }
        });
        break;
      }

      case "text": {
        // (str, x, y, w, h, opts)
        // opts: { fontSize, bold, color, align, valign, italic }
        const opts = a[5] || {};
        const runs = parseInlineBold(a[0]||"", opts.fontSize||11, resolveColor(opts.color)||C.dark);
        slide.addText(runs, {
          x: a[1]||0, y: a[2]||0, w: a[3]||2, h: a[4]||0.4,
          fontSize:  opts.fontSize  || 11,
          fontFace:  FONT,
          color:     resolveColor(opts.color) || C.dark,
          bold:      opts.bold    || false,
          italic:    opts.italic  || false,
          align:     opts.align   || "left",
          valign:    opts.valign  || "top",
          margin: 0
        });
        break;
      }

      case "watermarkX":
        // (accentKey)
        addGhostedX(slide, a[0] || "purple");
        break;

      default:
        console.warn("custom_diagram: unknown primitive fn: " + d.fn);
    }
  });
}

// BRIEF — EDIT THIS SECTION TO PRODUCE A NEW DECK
// ============================================================

const BRIEF = {
  title: 'CG Generator — Master Test Deck',
  client: 'Collaborate Middle East',
  slides: [
    { type: 'title', label: 'Internal Reference Deck', date: 'March 2026', headline: 'CG Generator\nMaster Test Deck' },
    { type: 'statement_of_intent', headline: 'Every slide type.\nOne deck.\nBuilt to impress.', imagePrompt: 'Bright, energetic creative studio environment. Vibrant gradient lighting in purple and blue. Wide angle, cinematic, no people, no text overlays. Generate in 16:9 widescreen landscape format.', googleSearch: 'creative studio vibrant colourful professional workspace' },
    { type: 'contents', items: [ { label: 'Text & Layout Slides', divider: 'Text & Layout' }, { label: 'Two & Three Column Layouts', divider: 'Two & Three Columns' }, { label: 'Charts — Bar, Line, Pie, Combo', divider: 'Charts' }, { label: 'Diagrams & Data Visualisation', divider: 'Diagrams' }, { label: 'Process, Strategy & Data', divider: 'Data Tables' }, { label: 'Image & Media Slides', divider: 'Image Slides' } ] },
    { type: 'divider', title: 'Text & Layout' },
    { type: 'headline_body', section: 'Overview', boldWord: 'Collaborate', rest: 'Global brand system', body: 'This master test deck exercises every slide type in the CG generator to validate output quality, layout consistency and brand compliance.' },
    { type: 'three_column', section: 'Capabilities', headline: 'Three areas of focus', headlineBold: 'for this year', columns: [ { body: ['[bold]Creative Excellence[/bold]\nPushing the boundaries of live experience with bold, original concepts.', 'Our creative team delivers cohesive brand worlds.'] }, { body: ['[bold]Data-Led Strategy[/bold]\nEvery recommendation is grounded in audience insight and competitive benchmarking.', 'We measure what matters, not just what is easy to count.'] }, { body: ['[bold]Flawless Delivery[/bold]\nWorld-class production capabilities across 40+ markets.', 'Logistics, staffing and technology handled end-to-end.'] } ] },
    { type: 'key_stats', section: 'Performance', headline: { normal: 'Results', bold: 'that speak for themselves' }, stats: [ { number: '£42M', label: 'Total client spend\nmanaged in 2025' }, { number: '97%', label: 'Client retention\nrate year-on-year' }, { number: '140+', label: 'Events delivered\nacross 38 markets' }, { number: '4.9', label: 'Average client\nsatisfaction score' } ] },
    { type: 'quote', section: 'Client Voice', quote: 'Collaborate Global consistently deliver work that exceeds our expectations. They understand our brand, challenge our thinking, and execute with a level of precision we have never experienced elsewhere.', attribution: '— Chief Marketing Officer, Global Automotive Brand' },
    { type: 'four_column', section: 'Services', headline: { normal: 'Our four core', bold: 'service lines' }, columns: [ { title: 'Brand Experience', body: 'Immersive physical and digital brand worlds. Launch events, fan activations, trade shows and owned IP.' }, { title: 'Employee Engagement', body: 'Conferences, awards ceremonies, town halls and incentive travel. Connecting people to purpose.' }, { title: 'Sponsorship Activation', body: 'Maximising the value of sports, music and cultural partnerships through creative activation programmes.' }, { title: 'Digital & Hybrid', body: 'Broadcast-quality virtual events, hybrid formats and always-on digital experience platforms.' } ] },
    { type: 'blue_cards', section: 'Process', boldWord: 'CG', rest: 'way of working', cards: [ { title: 'Discover', body: 'Deep-dive briefing sessions, stakeholder interviews and category analysis.', additional: ['Audience research', 'Competitor mapping', 'Brand audit'] }, { title: 'Define', body: 'Strategic platform and creative territories. Clear success metrics agreed before any concept work begins.', additional: ['Strategic brief', 'KPI framework', 'Budget parameters'] }, { title: 'Design', body: 'Concept development, creative direction, visual world-building and experience design.', additional: ['Concept decks', 'Mood boards', 'Experience journey'] }, { title: 'Deliver', body: 'End-to-end production management. Site, supplier, logistics, staffing and on-the-day execution.', additional: ['Production schedule', 'Live management', 'Post-event report'] } ] },
    { type: 'challenge', section: 'Case Study', headline: 'Global Auto\nLaunch', sections: [ { label: 'Challenge', body: 'Create a global product launch event for a flagship EV model, simultaneously in 12 cities across 3 continents.' }, { label: 'Idea', body: 'A single connected live moment broadcast across all venues — one stage, one story, delivered everywhere at the same second.' }, { label: 'Experience', body: 'Simultaneous reveal in 12 cities. 6,200 guests. 240M earned media impressions.' } ] },
    { type: 'timeline', section: 'Project Plan', title: 'Delivery timeline', rowLabels: ['CREATIVE', 'PRODUCTION'], phases: [ { label: 'STRATEGY', startX: 1.8, endX: 4.5 }, { label: 'CONCEPT', startX: 4.5, endX: 7.2 }, { label: 'BUILD', startX: 7.2, endX: 10.5 }, { label: 'LIVE', startX: 10.5, endX: 13.04 } ], topEvents: [ { x: 2.5, date: 'Wk 1', label: 'Briefing & kick-off' }, { x: 4.2, date: 'Wk 3', label: 'Creative territories' }, { x: 6.2, date: 'Wk 5', label: 'Concept sign-off' }, { x: 8.5, date: 'Wk 9', label: 'Creative production' }, { x: 11.5, date: 'Wk 14', label: 'Final rehearsal' } ], bottomEvents: [ { x: 3.2, date: 'Wk 2', label: 'Venue confirmed' }, { x: 5.5, date: 'Wk 4', label: 'Supplier briefings' }, { x: 7.8, date: 'Wk 7', label: 'Tech spec locked' }, { x: 10.0, date: 'Wk 11', label: 'Build commences' }, { x: 12.5, date: 'Wk 16', label: 'Event live' } ] },
    { type: 'team_cards', section: 'The Team', boldWord: 'core team', members: [ { name: 'Sarah Mitchell', title: 'Account Director', bio: 'Ten years in experiential marketing across luxury and automotive sectors. Sarah leads client relationships from brief to delivery with precision and warmth.' }, { name: 'James Okafor', title: 'Creative Director', bio: 'Award-winning creative with a background in theatre and brand design. James brings bold conceptual thinking to every project he leads.' }, { name: 'Priya Sharma', title: 'Senior Producer', bio: 'Specialist in large-scale multi-market events across APAC and EMEA. Priya is the engine behind flawless on-site delivery.' }, { name: 'Tom Reeves', title: 'Strategy Lead', bio: 'Strategist with deep expertise in audience insight and campaign measurement. Tom ensures every decision is grounded in data.' } ] },
    { type: 'divider', title: 'Two & Three Columns' },
    { type: 'two_column_text', section: 'Approach', boldWord: 'Two-column', rest: 'layout example', columns: [ { label: 'Left Column', accent: 'purple', body: '[bold]Use for two equally weighted ideas.[/bold]\n\nTwo-column layouts suit comparisons, before/after structures, challenge and response. The 1cm central gutter gives each column breathing room.' }, { label: 'Right Column', accent: 'green', body: '[bold]Columns do not need equal length.[/bold]\n\nShorter right columns often create a more dynamic layout than forcing symmetry. Accent colours cycle through the brand palette.' } ] },
    { type: 'three_column_text', section: 'Methodology', boldWord: 'Three-column', rest: 'layout example', columns: [ { label: 'Column One', accent: 'purple', body: '[bold]Three parallel ideas.[/bold]\n\nThree-column layouts suit roadmaps, stage-based processes and three-part frameworks where ideas carry equal weight.' }, { label: 'Column Two', accent: 'blue', body: '[bold]Keep body copy concise.[/bold]\n\nAt three columns the readable width per column is narrower. Use [bold]inline bold[/bold] to anchor key points.' }, { label: 'Column Three', accent: 'orange', body: '[bold]Accent colours are independent.[/bold]\n\nEach column takes its own accent colour. If no accent is specified the generator cycles automatically.' } ] },
    { type: 'divider', title: 'Charts' },
    { type: 'bar_chart', section: 'DATA', boldWord: 'Event attendance', rest: 'by sector — 2024 vs 2025', labels: ['Technology', 'Automotive', 'Sports', 'Finance', 'Retail'], series: [ { name: '2024', values: [1200, 850, 2100, 600, 950] }, { name: '2025', values: [1800, 1100, 2400, 750, 1300] } ], commentary: 'Attendance grew across all sectors in 2025. Technology showed the strongest proportional growth at 50% year-on-year.', insight: 'Technology leads growth at +50% YoY — strongest performer across all sectors.' },
    { type: 'horiz_bar_chart', section: 'DATA', boldWord: 'Budget allocation', rest: 'by discipline', labels: ['Production', 'Creative', 'Digital', 'Logistics', 'Staffing'], series: [{ name: 'Budget %', values: [38, 22, 18, 12, 10] }], showValue: true, commentary: 'Production remains the largest cost centre at 38%. Digital allocation at 18% reflects increased investment in hybrid platforms.', insight: 'Production at 38% is consistent with industry benchmarks for large-scale experiential.' },
    { type: 'line_chart', section: 'DATA', boldWord: 'Revenue growth', rest: '3-year trend vs target', labels: ['Q1 23','Q2 23','Q3 23','Q4 23','Q1 24','Q2 24','Q3 24','Q4 24','Q1 25'], series: [ { name: 'Actual £M', values: [1.2, 1.4, 1.8, 2.1, 2.0, 2.3, 2.7, 3.1, 2.9] }, { name: 'Target £M', values: [1.5, 1.5, 1.8, 2.0, 2.2, 2.4, 2.6, 2.8, 3.0] } ], commentary: 'Revenue tracked above target for 6 of the last 9 quarters. The Q1 2025 dip reflects seasonal pipeline timing.', insight: 'Tracked above target in 6 of 9 quarters — strong underlying trajectory.' },
    { type: 'pie_chart', section: 'DATA', boldWord: 'Revenue', rest: 'by client type — 2025', chartData: [ { label: 'Retained clients', value: 58 }, { label: 'New business', value: 27 }, { label: 'Agency partners', value: 15 } ], commentary: 'Retained client revenue at 58% reflects the strength of long-term relationships across our portfolio.', insight: 'Retained revenue majority confirms the strength of our long-term client relationships.' },
    { type: 'doughnut_chart', section: 'DATA', boldWord: 'Time allocation', rest: 'creative team — 2025', chartData: [ { label: 'Billable', value: 64 }, { label: 'Non-Billable', value: 23 }, { label: 'Internal', value: 13 } ], commentary: 'Billable utilisation at 64% is ahead of the 60% target. The 23% non-billable covers pitching, training and R&D.', insight: 'Billable utilisation at 64% — ahead of the 60% target.' },
    { type: 'combo_chart', section: 'DATA', boldWord: 'Pitches submitted', rest: 'vs win rate — H1 2025', labels: ['Jan','Feb','Mar','Apr','May','Jun'], barSeries: [{ name: 'Pitches submitted', values: [4, 6, 5, 8, 7, 9] }], lineSeries: [{ name: 'Win rate %', values: [25, 33, 40, 38, 43, 44] }], commentary: 'Win rate improved from 25% to 44% as pitch volume increased — demonstrating improved conversion quality.', insight: 'Win rate nearly doubled in 6 months — conversion quality improving alongside volume.' },
    { type: 'divider', title: 'Diagrams' },
    { type: 'flow_diagram', section: 'Process', boldWord: 'The intelligence', rest: 'engine — how it works', nodes: [ { title: 'Brief In', text: 'Client brief received, parsed and distributed to strategy and creative leads' }, { title: 'Research', text: 'Category analysis, audience mapping and competitive landscape review' }, { title: 'Strategy', text: 'Platform articulation, creative territories and KPI framework defined' }, { title: 'Concept', text: 'Concept development, mood boards, experience design and budget modelling' }, { title: 'Production', text: 'Supplier briefings, technical specification, schedule and risk register' }, { title: 'Delivery', text: 'Live event execution, real-time management and post-event reporting' } ] },
    { type: 'process_flow', section: 'Methodology', boldWord: 'Six-stage', rest: 'delivery process', steps: [ { date: 'Week 1', label: 'Discovery & brief', sub: 'Strategy', phase: 'DISCOVER' }, { date: 'Week 2', label: 'Research & insight', sub: 'Data', phase: 'DEFINE' }, { date: 'Week 4', label: 'Creative concept', sub: 'Creative', phase: 'DESIGN' }, { date: 'Week 6', label: 'Client sign-off', sub: 'Approval', phase: 'APPROVE' }, { date: 'Week 8', label: 'Production & build', sub: 'Production', phase: 'BUILD' }, { date: 'Week 12', label: 'Live event & reporting', sub: 'Delivery', phase: 'DELIVER' } ] },
    { type: 'convergence', section: 'Strategy', boldWord: 'Three forces', rest: 'driving our growth', items: [ { rank: 'Client demand', value: '+34%', label: 'YoY brief volume increase' }, { rank: 'Talent depth', value: '120+', label: 'Specialist staff globally' }, { rank: 'Market position', value: 'Top 3', label: 'UK experiential agencies' } ], convergence: { title: 'Sustainable scale', body: 'Combining demand momentum, talent infrastructure and market leadership to grow profitably without compromising quality.' } },
    { type: 'venn', section: 'Positioning', boldWord: 'Where', rest: 'we operate best', circles: [ { label: 'Brand', body: 'Deep brand purpose and positioning knowledge' }, { label: 'Audience', body: 'Behavioural insight across demographics' }, { label: 'Craft', body: 'Production excellence at every touchpoint' } ], outerLabel: 'The CG Advantage', subs: [ { label: 'Brand + Audience', body: 'Culturally relevant work that resonates — not just work that looks good.' }, { label: 'Audience + Craft', body: 'Experiences designed around real human behaviour, not assumed preferences.' }, { label: 'Brand + Craft', body: 'Production quality that elevates and protects the brand at every moment.' } ] },
    { type: 'concentric', section: 'Framework', boldWord: 'Brand', rest: 'experience hierarchy', circles: [ { title: 'Purpose', body: 'Why the brand exists and the role it plays in peoples lives' }, { title: 'Promise', body: 'What the brand commits to delivering in every interaction' }, { title: 'Presence', body: 'How the brand shows up physically, digitally and culturally' } ], notes: [ 'Purpose anchors all strategic decisions', 'Promise defines the experience standard', 'Presence is where production quality matters most' ] },
    { type: 'strategy_pillars', section: 'Strategy', boldWord: 'Growth', rest: 'strategy 2026', platform: 'CG Growth Platform: Deepen, Expand, Innovate', pillars: [ { title: 'Deepen', body: 'Grow wallet share within existing top-20 client accounts through proactive ideation and expanded service scope.' }, { title: 'Expand', body: 'Enter 6 new markets by end of 2026. Priority: APAC, MENA, DACH.' }, { title: 'Innovate', body: 'Launch CG Labs — exploring AI, spatial computing and new experience formats for 2027.' } ] },
    { type: 'matrix', section: 'Evaluation', boldWord: 'Capability', rest: 'vs market comparison', headers: ['Strategy', 'Creative', 'Production', 'Digital'], rows: [ { label: 'CG', cells: ['Market-leading', 'Award-winning', 'Global scale', 'Rapidly growing'] }, { label: 'Competitor A', cells: ['Strong', 'Strong', 'Regional only', 'Limited'] }, { label: 'Competitor B', cells: ['Limited', 'Market-leading', 'Strong', 'Market-leading'] }, { label: 'Competitor C', cells: ['Strong', 'Moderate', 'Strong', 'Moderate'] } ] },
    { type: 'divider', title: 'Data Tables' },
    { type: 'data_table', section: 'Financial Model', boldWord: 'Media Studio', rest: 'ROI summary', accent: 'purple', labelHeader: 'Metric', labelColW: 3.8, headers: ['Value'], rows: [ { label: 'Initial Investment', cells: ['£115,000'] }, { label: 'Annual Operating Costs', cells: ['£30,000'] }, { label: 'Annual Revenue', cells: ['£120,960'] }, { label: 'Annual Net Profit', cells: ['£90,960'], isBold: true }, { label: 'ROI', cells: ['79.1%'], isBold: true }, { label: 'Break-even Point', cells: ['1.26 years'], isBold: true } ], statsPanel: [ { value: '£90,960', label: 'Annual Net Profit' }, { value: '79.1%', label: 'Return on Investment' }, { value: '1.26 yrs', label: 'Break-even Point' }, { value: '£120,960', label: 'Total Annual Revenue' } ] },
    { type: 'custom_diagram', background: 'white', section: 'Custom Diagram', boldWord: 'Before', rest: 'and after — custom diagram example', draw: [ { fn: 'filledPill', args: ['Current State', 0.8, 1.85, 5.3, 0.36, 'cg', 11] }, { fn: 'filledPill', args: ['Future State', 7.3, 1.85, 5.3, 0.36, 'purple', 11] }, { fn: 'text', args: ['→', 6.1, 3.3, 1.1, 0.7, { fontSize: 36, bold: true, color: 'purple', align: 'center', valign: 'middle' }] }, { fn: 'roundRect', args: [0.8, 2.45, 5.3, 0.6, 'cg', 93, 0.06] }, { fn: 'text', args: ['Fragmented brief intake — no standard process', 0.95, 2.53, 5.0, 0.45, { fontSize: 10.5, valign: 'middle' }] }, { fn: 'roundRect', args: [0.8, 3.15, 5.3, 0.6, 'cg', 93, 0.06] }, { fn: 'text', args: ['Manual status tracking across disconnected tools', 0.95, 3.23, 5.0, 0.45, { fontSize: 10.5, valign: 'middle' }] }, { fn: 'roundRect', args: [7.3, 2.45, 5.3, 0.6, 'purple', 90, 0.06] }, { fn: 'text', args: ['Unified brief intake with AI-assisted categorisation', 7.45, 2.53, 5.0, 0.45, { fontSize: 10.5, valign: 'middle' }] }, { fn: 'roundRect', args: [7.3, 3.15, 5.3, 0.6, 'purple', 90, 0.06] }, { fn: 'text', args: ['Live project dashboard — one source of truth', 7.45, 3.23, 5.0, 0.45, { fontSize: 10.5, valign: 'middle' }] } ] },
    { type: 'divider', title: 'Image Slides' },
    { type: 'image_panel', section: 'Work', boldWord: 'Award-winning', rest: 'brand experience', body: 'Our 2025 campaign for a global technology brand won Best Brand Experience at the Event Industry Awards — attended by 12,000 consumers across 3 continents.', quote: 'The most impactful brand moment we have created in a decade.', imagePrompt: 'Large-scale brand activation event at night, dramatic stage lighting in blue and purple, thousands of engaged attendees, premium production quality. Editorial wide angle, no text overlays. Generate in 4:3 landscape format.', googleSearch: 'brand activation event stage lighting crowd night premium editorial' },
    { type: 'single_image', section: 'Production', boldWord: 'Precision', rest: 'at every scale', body: 'From intimate VIP dinners for 20 to stadium-filling brand moments for 50,000 — our production team delivers with the same precision and care.', quote: 'Not a single detail missed, at any scale.', imagePrompt: 'Professional event production crew at work — lighting rigs, technical equipment, backstage preparation. Behind the scenes, industrial aesthetic, dramatic lighting. No text overlays. Generate in 4:3 landscape format.', googleSearch: 'event production backstage technical crew lighting rigs industrial' },
    { type: 'full_bleed_text', section: 'Vision', headline: 'The future of live', boldWord: 'future', body: 'Immersive, intelligent, unforgettable experiences — delivered at scale across every market.', imagePrompt: 'Futuristic immersive brand event, laser lighting, projection mapping on architecture, small figures of attendees for scale. Cinematic wide angle, no text overlays. Generate in 16:9 widescreen landscape format.', googleSearch: 'projection mapping architecture event night immersive cinematic wide angle' },
    { type: 'two_images', section: 'Portfolio', boldWord: 'Two', rest: 'recent highlights', body: 'Selected work from our 2025 portfolio — representing the breadth of what we deliver across markets and disciplines.', quote: 'Consistency at scale is our differentiator.', imagePrompt1: 'Corporate awards ceremony — elegant black tie event, stage with dramatic lighting, seated audience. Premium, editorial quality. No text overlays. Generate in 4:3 landscape format.', googleSearch1: 'awards ceremony gala stage dramatic lighting black tie premium', imagePrompt2: 'Consumer brand activation — outdoor festival environment, vibrant crowd, colourful brand installation. Energetic, daylight. No text overlays. Generate in 4:3 landscape format.', googleSearch2: 'outdoor brand activation festival crowd colourful installation daylight' },
    { type: 'four_image_grid', section: 'Work', headline: '2025 highlight reel', images: [ { ratio: '4:3', prompt: 'High-end car reveal event, dramatic lighting, luxury venue interior, no people, premium atmosphere. Generate in 4:3 landscape format.' }, { ratio: '4:3', prompt: 'Consumer brand activation, vibrant crowd, outdoor festival, colourful branded installation. Generate in 4:3 landscape format.' }, { ratio: '4:3', prompt: 'Elegant black tie awards ceremony, stage with screen backdrop, audience silhouettes. Generate in 4:3 landscape format.' }, { ratio: '4:3', prompt: 'Tech product launch event, minimalist stage, blue lighting, product reveal moment. Generate in 4:3 landscape format.' } ] },
    { type: 'mood_board', images: [ { ratio: '2:3', prompt: 'Dramatic editorial brand portrait, rich jewel-toned colour, high contrast studio lighting. Generate in 2:3 portrait format.' }, { ratio: '4:3', prompt: 'Event crowd aerial wide shot, dramatic colourful lighting, large-scale production. Generate in 4:3 landscape format.' }, { ratio: '4:3', prompt: 'Luxury product detail shot, dark background, dramatic studio lighting, premium feel. Generate in 4:3 landscape format.' }, { ratio: '4:3', prompt: 'Creative light installation, geometric, minimal, architectural, blue and purple tones. Generate in 4:3 landscape format.' }, { ratio: '4:3', prompt: 'Speaker silhouette on stage against large illuminated screen, dramatic, cinematic. Generate in 4:3 landscape format.' } ] },
    { type: 'case_study', section: 'Case Study', boldWord: 'Nike', rest: 'Global Just Do It Tour', sections: [ { label: 'Challenge', body: 'Bring the Just Do It ethos to life in 10 cities simultaneously, connecting with local sporting communities authentically.' }, { label: 'Solution', body: 'City-specific activations anchored by a shared live broadcast moment — local heroes, global story.' }, { label: 'Result', body: '180,000 direct participants. 420M impressions. 94% brand sentiment lift across all markets.' } ], conclusion: 'Awarded Best Global Campaign at the Experiential Marketing Summit 2025.', images: [ { ratio: '4:3', prompt: 'Urban brand activation space, street-level, energetic crowd of young athletes, outdoor. Generate in 4:3 landscape format.' }, { ratio: '4:3', prompt: 'Diverse group of young athletes at outdoor event, energetic movement, natural light. Generate in 4:3 landscape format.' }, { ratio: '4:3', prompt: 'Large-scale creative brand installation in public space, geometric, bold branding. Generate in 4:3 landscape format.' }, { ratio: '4:3', prompt: 'Keynote reveal moment on stage, athlete on stage, dramatic uplighting, crowd reaction. Generate in 4:3 landscape format.' } ] },
    { type: 'case_study_hero', boldWord: 'Global', headline: 'Global Auto Launch 2025', stats: ['6,200 guests across 12 cities', '240M earned media impressions', '94% brand sentiment lift'], imagePrompt: 'Dramatic hero shot of a global automotive launch event, wide angle, cinematic, premium, dark moody atmosphere. Car silhouette on stage, dramatic lighting, no text overlays. Generate in 16:9 widescreen landscape format.' },
    { type: 'device_mockup', section: 'Digital', boldWord: 'App experience', rest: 'companion platform', body: 'Our event companion app delivered real-time personalisation, live polling and exclusive content to 40,000 attendees.', imagePrompt: 'Event companion mobile app UI, dark design, live schedule and content feed, branded purple and white, portrait orientation. Generate in 9:16 portrait format.' },
    { type: 'device_mood_board', images: [ { prompt: 'Event companion app home screen, dark UI, branded purple and white, portrait. Generate in 9:16 portrait format.' }, { prompt: 'Live event schedule and notifications screen, clean dark design, portrait. Generate in 9:16 portrait format.' }, { prompt: 'Interactive poll and engagement screen, vibrant, modern UI, portrait. Generate in 9:16 portrait format.' } ] },
    { type: 'ipad_video', section: 'Digital', boldWord: 'Digital content', rest: 'platform showcase', caption: 'The CG content platform — accessible to all attendees via tablet and mobile throughout the event.' },
    { type: 'phone_landscape', section: 'Digital', boldWord: 'Social content', rest: 'optimised for mobile', caption: 'All event content is formatted and delivered natively for mobile — landscape highlights and real-time social feeds.' },
    { type: 'full_bleed_caption', section: 'Work', headline: 'Where brands\ncome alive', subcaption: 'Collaborate Global — London, 2025', imagePrompt: 'Dramatic wide-angle photograph of a large-scale brand activation event. Thousands of people, spectacular lighting, premium production. Dark moody atmosphere, editorial quality. No text overlays. Generate in 16:9 widescreen landscape format.', googleSearch: 'brand activation large scale event dramatic lighting crowd night spectacular' },
    { type: 'sub_divider', title: 'Appendix' },
    { type: 'appendix_divider', title: 'Supporting Data' },
    { type: 'end' }
  ]
};

// ============================================================
// RUNNER — do not edit below this line

async function buildDeck(brief) {
  let pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";
  pres.author = "Collaborate Middle East";
  pres.title  = brief.title;

  for (const slide of brief.slides) {
    switch (slide.type) {
      case "title":          addTitleSlide(pres, slide);         break;
      case "statement_of_intent": addStatementOfIntentSlide(pres, slide); break;
      case "custom_diagram":   addCustomDiagramSlide(pres, slide);  break;
      case "custom_diagram":       addCustomDiagramSlide(pres, slide);       break;
      case "full_bleed":     addFullBleedSlide(pres, slide);     break;
      case "contents":       addContentsSlide(pres, slide);      break;
      case "divider":        addDividerSlide(pres, slide);       break;
      case "sub_divider":    addSubDividerSlide(pres, slide);    break;
      case "headline_body":  addHeadlineBodySlide(pres, slide);  break;
      case "three_column":   addThreeColumnSlide(pres, slide);   break;
      case "key_stats":      addKeyStatsSlide(pres, slide);      break;
      case "four_column":    addFourColumnSlide(pres, slide);    break;
      case "challenge":      addChallengeSlide(pres, slide);     break;
      case "single_image":   addSingleImageSlide(pres, slide);   break;
      case "two_images":     addTwoImagesSlide(pres, slide);     break;
      case "quote":          addQuoteSlide(pres, slide);         break;
      case "blue_cards":     addBlueCardsSlide(pres, slide);     break;
      case "timeline":       addTimelineSlide(pres, slide);      break;
      case "team_cards":     addTeamCardsSlide(pres, slide);     break;
      case "flow_diagram":   addFlowDiagramSlide(pres, slide);   break;
      case "end":            addEndSlide(pres);                  break;
      case "image_panel":    addImagePanelSlide(pres, slide);    break;
      case "full_bleed_text":addFullBleedTextSlide(pres, slide); break;
      case "four_image_grid":addFourImageGridSlide(pres, slide); break;
      case "mood_board":     addMoodBoardSlide(pres, slide);     break;
      case "case_study":     addCaseStudySlide(pres, slide);     break;
      case "case_study_hero":addCaseStudyHeroSlide(pres, slide); break;
      case "device_mockup":  addDeviceMockupSlide(pres, slide);  break;
      case "device_mood_board": addDeviceMoodBoardSlide(pres, slide); break;
      case "concentric":     addConcentricSlide(pres, slide);    break;
      case "venn":           addVennSlide(pres, slide);          break;
      case "convergence":    addConvergenceSlide(pres, slide);   break;
      case "process_flow":   addProcessFlowSlide(pres, slide);   break;
      case "strategy_pillars":addStrategyPillarsSlide(pres, slide); break;
      case "matrix":         addMatrixSlide(pres, slide);        break;
      case "bar_chart":      addBarChartSlide(pres, slide);      break;
      case "horiz_bar_chart":addHorizBarChartSlide(pres, slide); break;
      case "line_chart":     addLineChartSlide(pres, slide);     break;
      case "pie_chart":      addPieChartSlide(pres, slide);      break;
      case "doughnut_chart": addDoughnutChartSlide(pres, slide); break;
      case "combo_chart":    addComboChartSlide(pres, slide);    break;
      case "ipad_video":     addIpadVideoSlide(pres, slide);     break;
      case "appendix_divider": addAppendixDividerSlide(pres, slide);        break;
      case "two_column_text":  addTwoColumnTextSlide(pres, slide);           break;
      case "three_column_text":addThreeColumnTextSlide(pres, slide);         break;
      case "data_table":       addDataTableSlide(pres, slide);               break;
      case "phone_landscape":  addPhoneLandscapeVideoSlide(pres, slide);     break;
      case "full_bleed_caption":addFullBleedCaptionSlide(pres, slide);       break;
      default:
        console.warn(`Unknown slide type: "${slide.type}" — skipped`);
    }
  }

  await pres.writeFile({ fileName: OUTPUT_PATH });
  console.log("Deck saved to:", OUTPUT_PATH);
}

// ── PAGE NUMBER HELPER ────────────────────────────────────────────────────────
// Called at the end of every slide function EXCEPT title/divider/end types.
// Adds the actual slide number as a static text box, bottom-right, aligned
// with the X icon. The number is computed from the BRIEF slide index at
// generation time — it matches what PowerPoint will display.
//
// NOTE: these are static numbers baked at generation time. If slides are
// manually reordered in PowerPoint afterwards, regenerate the deck to refresh.

// PAGE NUMBERS: removed globally per design review.
// Numbers are not baked into slides — insert manually in PowerPoint if needed.

async function buildDeck(brief) {
  let pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";
  pres.author = "Collaborate Middle East";
  pres.title  = brief.title;

  // Compute contents page refs — for each contents item, find its matching divider
  // in the BRIEF (by matching item.label to divider title, case-insensitive), then
  // find the page number of the FIRST slide after that divider.
  // Falls back to sequential numbers if no match found.
  const DIVIDER_TYPES = new Set(["divider", "sub_divider", "appendix_divider"]);
  brief.slides.forEach((s, i) => {
    if (s.type === "contents" && s.items) {
      // Build a map: divider title (lowercase) -> page number of first slide after it
      const dividerPageMap = {};
      for (let d = 0; d < brief.slides.length; d++) {
        if (DIVIDER_TYPES.has(brief.slides[d].type)) {
          // Point at the divider slide itself (d + 1 is 1-based slide number)
          dividerPageMap[(brief.slides[d].title || "").toLowerCase().trim()] = d + 1;
        }
      }
      // Match each contents item to a divider
      // Priority: 1. item.divider exact match  2. label partial match  3. null (sequential fallback)
      s.pageRefs = s.items.map(item => {
        // Explicit divider key — most reliable
        if (item.divider) {
          const key = item.divider.toLowerCase().trim();
          if (dividerPageMap[key] !== undefined) return dividerPageMap[key];
        }
        // Label partial match fallback
        const key = (item.label || "").toLowerCase().trim();
        if (dividerPageMap[key] !== undefined) return dividerPageMap[key];
        for (const [dkey, dnum] of Object.entries(dividerPageMap)) {
          if (key.includes(dkey) || dkey.includes(key)) return dnum;
        }
        return null;
      });
      // Write pageRef back onto each item so addContentsSlide can use it for hyperlinks
      s.items.forEach((item, i) => { item.pageRef = s.pageRefs[i]; });
      console.log("Contents page refs:", s.items.map((item, i) => `${item.label} → ${s.pageRefs[i]}`).join(", "));
    }
  });

  for (let i = 0; i < brief.slides.length; i++) {
    const slide = brief.slides[i];

    switch (slide.type) {
      case "title":          addTitleSlide(pres, slide);         break;
      case "statement_of_intent": addStatementOfIntentSlide(pres, slide); break;
      case "custom_diagram":   addCustomDiagramSlide(pres, slide);  break;
      case "custom_diagram":       addCustomDiagramSlide(pres, slide);       break;
      case "full_bleed":     addFullBleedSlide(pres, slide);     break;
      case "contents":       addContentsSlide(pres, slide);      break;
      case "divider":        addDividerSlide(pres, slide);       break;
      case "sub_divider":    addSubDividerSlide(pres, slide);    break;
      case "headline_body":  addHeadlineBodySlide(pres, slide);  break;
      case "three_column":   addThreeColumnSlide(pres, slide);   break;
      case "key_stats":      addKeyStatsSlide(pres, slide);      break;
      case "four_column":    addFourColumnSlide(pres, slide);    break;
      case "challenge":      addChallengeSlide(pres, slide);     break;
      case "single_image":   addSingleImageSlide(pres, slide);   break;
      case "two_images":     addTwoImagesSlide(pres, slide);     break;
      case "quote":          addQuoteSlide(pres, slide);         break;
      case "blue_cards":     addBlueCardsSlide(pres, slide);     break;
      case "timeline":       addTimelineSlide(pres, slide);      break;
      case "team_cards":     addTeamCardsSlide(pres, slide);     break;
      case "flow_diagram":   addFlowDiagramSlide(pres, slide);   break;
      case "end":            addEndSlide(pres);                  break;
      case "image_panel":    addImagePanelSlide(pres, slide);    break;
      case "full_bleed_text":addFullBleedTextSlide(pres, slide); break;
      case "four_image_grid":addFourImageGridSlide(pres, slide); break;
      case "mood_board":     addMoodBoardSlide(pres, slide);     break;
      case "case_study":     addCaseStudySlide(pres, slide);     break;
      case "case_study_hero":addCaseStudyHeroSlide(pres, slide); break;
      case "device_mockup":  addDeviceMockupSlide(pres, slide);  break;
      case "device_mood_board": addDeviceMoodBoardSlide(pres, slide); break;
      case "concentric":     addConcentricSlide(pres, slide);    break;
      case "venn":           addVennSlide(pres, slide);          break;
      case "convergence":    addConvergenceSlide(pres, slide);   break;
      case "process_flow":   addProcessFlowSlide(pres, slide);   break;
      case "strategy_pillars":addStrategyPillarsSlide(pres, slide); break;
      case "matrix":         addMatrixSlide(pres, slide);        break;
      case "bar_chart":      addBarChartSlide(pres, slide);      break;
      case "horiz_bar_chart":addHorizBarChartSlide(pres, slide); break;
      case "line_chart":     addLineChartSlide(pres, slide);     break;
      case "pie_chart":      addPieChartSlide(pres, slide);      break;
      case "doughnut_chart": addDoughnutChartSlide(pres, slide); break;
      case "combo_chart":    addComboChartSlide(pres, slide);    break;
      case "ipad_video":     addIpadVideoSlide(pres, slide);     break;
      case "appendix_divider": addAppendixDividerSlide(pres, slide); break;
      case "two_column_text":  addTwoColumnTextSlide(pres, slide);   break;
      case "three_column_text":addThreeColumnTextSlide(pres, slide); break;
      case "data_table":       addDataTableSlide(pres, slide);       break;
      case "phone_landscape":  addPhoneLandscapeVideoSlide(pres, slide); break;
      case "full_bleed_caption":addFullBleedCaptionSlide(pres, slide); break;
      default:
        console.warn(`Unknown slide type: "${slide.type}" — skipped`);
    }

    // Page numbers removed — no static numbers baked into slides
  }

  await pres.writeFile({ fileName: OUTPUT_PATH });
  console.log(`Deck saved to: ${OUTPUT_PATH}`);
}

buildDeck(BRIEF);
