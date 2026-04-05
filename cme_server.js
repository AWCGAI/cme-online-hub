/**
 * CME Online Hub — Express Server
 * Collaborate Middle East PowerPoint Generator — Web Deployment
 *
 * POST /generate   → Accepts files + text content, returns .pptx download
 * GET  /health     → Health check for Railway
 */

const express    = require('express');
const path       = require('path');
const fs         = require('fs');
const { execFile } = require('child_process');
const { v4: uuidv4 } = require('uuid');
const Anthropic  = require('@anthropic-ai/sdk');
const multer     = require('multer');
const mammoth    = require('mammoth');
const XLSX       = require('xlsx');
const pdfParse   = require('pdf-parse');

const app  = express();
const PORT = process.env.PORT || 3000;

// ── PATHS ─────────────────────────────────────────────────────────────────────
const GENERATOR_PATH = path.join(__dirname, 'cme_generator.js');
const ASSETS_DIR     = path.join(__dirname, 'cme_assets');

// ── CORS ──────────────────────────────────────────────────────────────────────
app.use((req, res, next) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, GET, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.sendStatus(200);
  next();
});

app.use(express.json({ limit: '2mb' }));
app.use(express.static(__dirname));

// ── MULTER (memory storage, 50MB limit) ───────────────────────────────────────
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 50 * 1024 * 1024 }
});

// ── HEALTH CHECK ──────────────────────────────────────────────────────────────
app.get('/health', (req, res) => {
  res.json({ status: 'ok', timestamp: new Date().toISOString() });
});

// ── DEBUG: return last generated BRIEF ────────────────────────────────────────
let _lastBrief = null;
app.get('/debug/brief', (req, res) => {
  if (!_lastBrief) return res.json({ brief: null, message: 'No generation yet this session' });
  res.type('text/plain').send(_lastBrief);
});

// ── FILE TEXT EXTRACTION ──────────────────────────────────────────────────────
async function extractFileText(file) {
  const ext = path.extname(file.originalname).toLowerCase();
  const name = file.originalname;

  try {
    // Word documents
    if (ext === '.docx' || ext === '.doc') {
      const result = await mammoth.extractRawText({ buffer: file.buffer });
      return `[Document: ${name}]\n${result.value}`;
    }

    // PDF
    if (ext === '.pdf') {
      const result = await pdfParse(file.buffer);
      return `[PDF: ${name}]\n${result.text}`;
    }

    // Excel / CSV
    if (ext === '.xlsx' || ext === '.xls' || ext === '.csv') {
      const wb = XLSX.read(file.buffer, { type: 'buffer' });
      const sheets = wb.SheetNames.map(sheetName => {
        const ws = wb.Sheets[sheetName];
        return `Sheet "${sheetName}":\n` + XLSX.utils.sheet_to_csv(ws);
      });
      return `[Spreadsheet: ${name}]\n${sheets.join('\n\n')}`;
    }

    // Plain text / Markdown
    if (ext === '.txt' || ext === '.md') {
      return `[Text file: ${name}]\n${file.buffer.toString('utf8')}`;
    }

    // PowerPoint — note it but can't extract without additional parser
    if (ext === '.pptx' || ext === '.ppt') {
      return `[PowerPoint: ${name} — uploaded as reference. Use the paste field to describe what to generate based on this presentation.]`;
    }

    // Images — note them for context
    if (['.jpg', '.jpeg', '.png', '.gif', '.webp'].includes(ext)) {
      return `[Image: ${name} — uploaded as visual reference.]`;
    }

    // Fallback: attempt UTF-8 text read
    return `[File: ${name}]\n${file.buffer.toString('utf8')}`;

  } catch (err) {
    return `[File: ${name} — could not extract text: ${err.message}]`;
  }
}

// ── SLIDE TYPE SCHEMA (injected into Haiku system prompt) ──────────────────────
const SLIDE_SCHEMA = `
AVAILABLE SLIDE TYPES AND REQUIRED FIELDS:

title: { type, label, date, headline (two lines split by \\n) }
contents: { type, section, items: [{ label, divider }] }
divider: { type, title }
sub_divider: { type, title }
end: { type }
headline_body: { type, section, boldWord, rest, body }
two_column_text: { type, section, headline, columns: [{ label, body }, { label, body }] }
three_column_text: { type, section, headline, columns: [{ label, body }, { label, body }, { label, body }] }
three_column: { type, section, headline, headlineBold, columns: [{ title, body }] }
four_column: { type, section, headline, columns: [{ title, body }] }
key_stats: { type, section, stats: [{ number, label }] }
quote: { type, quote, attribution }
blue_cards: { type, section, headline, cards: [{ title, body }] }
challenge: { type, section, headline, headlineBold, sections: [{ label, body }] }
flow_diagram: { type, section, boldWord, rest, nodes: [{ title, text }] }
process_flow: { type, section, boldWord, rest, steps: [{ title, body }] }
strategy_pillars: { type, section, boldWord, rest, platform, pillars: [{ title, body }] }
convergence: { type, section, boldWord, rest, forces: [{ rank, label, body }], outcome: { label, body } }
venn: { type, section, boldWord, rest, circles: [{ label, body }], overlap }
concentric: { type, section, boldWord, rest, rings: [{ label, body }] }
matrix: { type, section, boldWord, rest, headers: [], rows: [[]] }
timeline: { type, section, boldWord, rest, lanes: [{ label, events: [{ label, date, body }] }] }
bar_chart: { type, section, headline, labels: [], series: [{ name, values: [] }] }
line_chart: { type, section, headline, labels: [], series: [{ name, values: [] }] }
pie_chart: { type, section, headline, chartData: [{ label, value }] }
doughnut_chart: { type, section, headline, chartData: [{ label, value }] }
horiz_bar_chart: { type, section, headline, labels: [], series: [{ name, values: [] }] }
combo_chart: { type, section, headline, labels: [], barSeries: [{ name, values: [] }], lineSeries: [{ name, values: [] }] }
image_panel: { type, section, boldWord, rest, body, imagePrompt, googleSearch }
single_image: { type, section, boldWord, rest, body, imagePrompt, googleSearch }
full_bleed_text: { type, section, headline, imagePrompt, googleSearch }
full_bleed_caption: { type, section, headline, subcaption, imagePrompt, googleSearch }
two_images: { type, section, boldWord, rest, body, images: [{ prompt, ratio }, { prompt, ratio }] }
four_image_grid: { type, section, boldWord, rest, images: [{ prompt, ratio }] }
mood_board: { type, imagePrompts: ['...', '...', '...', '...', '...'] }
case_study: { type, section, boldWord, rest, sections: [{ label, body }], conclusion, images: [{ ratio, prompt }] }
case_study_hero: { type, boldWord, headline, stats: [], imagePrompt }
team_cards: { type, section, headline, members: [{ name, title, boldWord }] }
data_table: { type, section, headline, headers: [], rows: [[]], statsPanel: [{ number, label }] }

SCHEMA RULES (these caused crashes in testing — follow exactly):
- strategy_pillars: platform must be a plain string, never an object
- team_cards: role field key is 'title' not 'role'. boldWord is bold part only (generator prepends 'Your ' automatically)
- two_column_text: requires columns array — NOT col1/col2 keys
- blue_cards: do NOT include 'The' in boldWord — generator prepends it
- three_column and challenge: use headline/headlineBold, NOT boldWord/rest
- key_stats: field is 'number' not 'value'
- flow_diagram nodes: use title/text, NOT label/detail
- contents items: divider key must exactly match the divider slide's title field
`.trim();

// ── HAIKU SYSTEM PROMPT ───────────────────────────────────────────────────────
const SYSTEM_PROMPT = `You are a specialist in creating structured content for the Collaborate Global PowerPoint Generator.
Your job is to convert free-form text into a valid JavaScript BRIEF object.

CRITICAL OUTPUT RULES:
- Respond ONLY with a JavaScript object literal — the value of the BRIEF constant
- No markdown code fences, no preamble, no explanation, no trailing text
- Use ONLY single-quoted strings — never double quotes, never smart/curly quotes
- Escape apostrophes with backslash: it\\'s, we\\'ve
- Use \\n for line breaks within strings (never literal newlines inside string values)
- Start your response with { and end with }
- Every image slide must include imagePrompt and googleSearch fields

BRIEF STRUCTURE:
{
  title: 'Deck Title',
  client: 'Client Name',
  slides: [ ...array of slide objects... ]
}

DESIGN RULES:
- Open with title slide, then contents, then dividers for each section, close with end
- Use 12-20 slides total for most content
- Vary slide types — no two consecutive slides of the same type
- Use diagram slides (flow_diagram, process_flow, venn, strategy_pillars) for processes/frameworks
- Use key_stats or chart slides whenever there are numbers
- Use image_panel or full_bleed_caption to break up sections
- contents items must have divider keys matching exactly the title of their section dividers
- Write in third person about Collaborate Middle East. Use British English.

${SLIDE_SCHEMA}


COLLABORATE MIDDLE EAST — COMPANY KNOWLEDGE:

# COLLABORATE MIDDLE EAST — KNOWLEDGE FILE
## For use in the CME PowerPoint Generator system prompt

---

## WHO WE ARE

Collaborate Middle East is the GCC-focused arm of Collaborate Global — an independent, award-winning global brand experience agency. We operate at the intersection of world-class international creativity and deep regional cultural intelligence, delivering visionary strategies and measurable impact across Saudi Arabia and the wider GCC.

We are headquartered in Riyadh, KSA with a second office in Dubai, UAE. Our parent network spans Europe (Chichester and London), Asia (Tokyo), Africa (Port Elizabeth), and the Americas (opening 2025).

Our founding belief: world-class creativity meets local ambition.

---

## WHAT WE DO

We provide experiential strategy, ideation, design and delivery across the GCC. Our offer combines creative excellence with production intelligence and a hyperservice mentality.

**Our full-service disciplines:**

- **Research & Strategy** — Research, data mining, behavioural science, culture, category, commerce, consumer, media, KPI definition, short and long-term goal-setting
- **Creative Ideation** — Positioning, conceptual development, transformative 'Long' ideas, activation platforms, brand expression, tech innovation
- **Campaign Design & Development** — Customer journeys, 3D design, 2D design, digital content, film, writing
- **Project Management & Production** — Procurement, purchasing, RFP management, full-service project management, in-house production, outsourced production, staffing services
- **Live Activations** — One-off brand activations, roadshows, semi-permanent installations, museums, fanzones, branded spaces, entertainment zones, brand theatre, heritage
- **Amplifications** — Social campaigns, PR, influencer, fanbase building, content studio
- **Measurement & Evaluation** — Engagement levels, reach, commercial success, design iteration, impact

---

## OUR THREE PILLARS

**1. Heritage, Culture & Vision**
We partner with organisations and leaders in Saudi Arabia to craft visionary communication strategies and long-term, experiential project plans, leveraging international expertise to drive measurable societal and community impact. All powered by a unique AI cultural tool.

**2. Partnerships, Sponsorships & Collaborations**
Our most valuable impact is delivered by bringing together both strands of our expertise. Society-driven, experiential moments require brands to partner, sponsor and collaborate. The intersection of both our worlds drives value and impact to the other.

**3. Culture-Led Brand Expressions**
We energise brands across Saudi Arabia and the GCC by combining international expertise with local cultural insights, delivering creative, high-impact strategies and campaigns that drive immediate results and long-term value.

---

## OUR POSITIONING

**Importing brands into the region, exporting brands out, and building local brands in market.**

- We bring international expertise in brand communication and long-term project planning to Saudi Arabia, delivering strategies aligned with Vision 2030.
- As trusted partners, we fuse global insights with local understanding to create measurable societal and community impact.
- Fresh thinking, agile execution and impactful results, tailored to the fast-paced GCC market.

---

## OUR VALUES

**Global Thinking, Regional Impact** — We adapt world-class expertise to create solutions that resonate with GCC audiences.

**Dynamic and Agile** — We thrive on energy, adaptability, and speed to deliver in a fast-paced market.

**Creativity with Purpose** — Our imaginative solutions are grounded in delivering real business results.

**Ambitious Collaboration** — We align with brands seeking bold ideas and transformative growth.

**ROI-Driven** — We prioritise measurable returns and outcomes for every project.

---

## OUR CLIENTS & PARTNERS

We have worked with global and regional brands including:

- **Al Madinah Region Development Authority** — Saudi government cultural infrastructure
- **Qatar Airways** — Regional aviation and hospitality
- **Aston Martin** — Premium automotive brand experience in the Gulf
- **Louis Vuitton** — Luxury fashion in the GCC
- **WPP** — Global marketing communications
- **Novartis** — Pharmaceutical brand engagement
- **Vita Coco** — Consumer brand activation
- **Hispano Suiza** — Ultra-premium automotive
- **Audi** — Saudi Arabia production campaign
- **Qiddiya** — Saudi entertainment city experience
- **Smart Madinah** — City visitor experience

---

## LEADERSHIP TEAM

**Ben McMahon** — Global CEO & Founder
**Nick Walsh** — Client Experience Director
**Andrew Walker** — Chief Creative Officer
**Natassha Evans** — Head of Client Experience
**Ewan Ferrier** — Creative Director
**Hamad Tariq Mahmoud** — Saudi Cultural Lead
**Hamsa Amjed** — Global Consultant
**Ross Oxenham** — Chief Expansion Officer
**Abdulrahman Alrashidi** — Agency Manager

---

## WRITING GUIDELINES FOR THE GENERATOR

- Write in third person about Collaborate Middle East (e.g. "Collaborate Middle East delivers..." not "We deliver...")
- Use British English spelling throughout
- Reference the GCC, Saudi Arabia, KSA, UAE, Riyadh, Dubai as appropriate to context
- Tone: confident, cultural, premium — not corporate or generic
- When referencing the global network, mention Collaborate Global as the parent
- Vision 2030 is a key context for Saudi Arabia work — reference where relevant
- The agency's unique differentiator is the fusion of international expertise with deep local cultural intelligence

`;

// ── BRIEF GENERATION VIA HAIKU ────────────────────────────────────────────────
async function generateBrief(userContent) {
  const client = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

  const response = await client.messages.create({
    model: 'claude-haiku-4-5-20251001',
    max_tokens: 4096,
    system: SYSTEM_PROMPT,
    messages: [{
      role: 'user',
      content: `Convert this content into a BRIEF object for the CME Generator:\n\n${userContent}`
    }]
  });

  let text = response.content[0].text.trim();

  // Strip any code fences if the model ignored instructions
  if (text.startsWith('```')) {
    text = text.replace(/^```[^\n]*\n?/, '').replace(/```\s*$/, '').trim();
  }

  // Strip any trailing explanation after the closing brace
  const lastBrace = text.lastIndexOf('}');
  if (lastBrace !== -1) {
    text = text.slice(0, lastBrace + 1);
  }

  // Evaluate to validate it's a real object
  let brief;
  try {
    // eslint-disable-next-line no-eval
    brief = eval('(' + text + ')');
  } catch (e) {
    throw new Error(`BRIEF evaluation failed: ${e.message}\n\nRaw AI response:\n${text}`);
  }

  if (!brief || !brief.slides || !Array.isArray(brief.slides)) {
    throw new Error('BRIEF object is missing slides array');
  }

  return { brief, briefText: text };
}

// ── BUILD .PPTX ───────────────────────────────────────────────────────────────
function buildPptx(briefText) {
  return new Promise((resolve, reject) => {
    const runId      = uuidv4();
    const runnerPath = `/tmp/run_${runId}.js`;
    const outputPath = `/tmp/cg_output_${runId}.pptx`;

    // Read the full generator source
    let source = fs.readFileSync(GENERATOR_PATH, 'utf8');

    // Patch hardcoded paths wherever they appear
    source = source
      .replace(
        'const ASSETS_DIR = "/home/claude/cme_assets"',
        `const ASSETS_DIR = ${JSON.stringify(ASSETS_DIR)}`
      )
      .replace(
        'const OUTPUT_PATH = "/home/claude/cme_output.pptx"',
        `const OUTPUT_PATH = ${JSON.stringify(outputPath)}`
      );

    // Replace the BRIEF object in-place so buildDeck (defined AFTER it) stays intact.
    // Splice point is the start of "async function buildDeck(" — NOT the buildDeck(BRIEF) call.
    // This preserves: [engine] + [new BRIEF] + [buildDeck definition] + [buildDeck(BRIEF) call]
    const briefStart       = source.indexOf('const BRIEF = ');
    const buildDeckDefStart = source.indexOf('\nasync function buildDeck(', briefStart);

    if (briefStart === -1) {
      return reject(new Error('Could not find "const BRIEF" in cme_generator.js'));
    }
    if (buildDeckDefStart === -1) {
      return reject(new Error('Could not find "async function buildDeck(" in cme_generator.js'));
    }

    // Stitch: everything before BRIEF + new BRIEF + buildDeck definition onward
    const runner =
      source.slice(0, briefStart) +
      `const BRIEF = ${briefText};\n` +
      source.slice(buildDeckDefStart + 1); // +1 skips the leading \n

    fs.writeFileSync(runnerPath, runner, 'utf8');

    execFile('node', [runnerPath], {
      cwd: __dirname,
      timeout: 120000,
      maxBuffer: 10 * 1024 * 1024,
      env: {
        ...process.env,
        NODE_PATH: path.join(__dirname, 'node_modules')
      }
    }, (err, stdout, stderr) => {
      // Always clean up runner
      try { fs.unlinkSync(runnerPath); } catch (_) {}

      if (err) {
        try { fs.unlinkSync(outputPath); } catch (_) {}
        return reject(new Error(stderr || stdout || err.message));
      }

      if (!fs.existsSync(outputPath)) {
        return reject(new Error('Generator ran but output file not found'));
      }

      resolve(outputPath);
    });
  });
}

// ── GENERATE ENDPOINT (multipart + JSON fallback) ─────────────────────────────
app.post('/generate', upload.array('files', 20), async (req, res) => {

  if (!process.env.ANTHROPIC_API_KEY) {
    return res.status(500).json({ error: 'ANTHROPIC_API_KEY environment variable is not set.' });
  }

  // Extract text from each uploaded file
  const fileTexts = [];
  if (req.files && req.files.length > 0) {
    for (const file of req.files) {
      const extracted = await extractFileText(file);
      fileTexts.push(extracted);
    }
  }

  // Combine with pasted content
  const pastedContent = (req.body.content || '').trim();
  const allContent = [...fileTexts, pastedContent].filter(Boolean).join('\n\n---\n\n');

  if (allContent.length < 10) {
    return res.status(400).json({ error: 'Please provide content — upload a file or paste some text.' });
  }

  console.log(`[${new Date().toISOString()}] Generating: ${req.files ? req.files.length : 0} file(s), ${pastedContent.length} chars pasted`);

  let briefText, outputPath;

  try {
    const result = await generateBrief(allContent);
    briefText = result.briefText;
    _lastBrief = briefText;   // store for /debug/brief
    console.log(`Brief generated: ${result.brief.slides.length} slides`);
  } catch (err) {
    console.error('Brief generation failed:', err.message);
    return res.status(500).json({ error: `Brief generation failed: ${err.message}` });
  }

  try {
    outputPath = await buildPptx(briefText);
    console.log(`PPTX built: ${outputPath}`);
  } catch (err) {
    console.error('PPTX build failed:', err.message);
    return res.status(500).json({ error: `PPTX build failed: ${err.message}` });
  }

  // Determine filename for download
  const requestedName = (req.body.filename || '').trim().replace(/[^a-zA-Z0-9_\-\.]/g, '_');
  const downloadName  = (requestedName || 'collaborate-middle-east-presentation') + '.pptx';

  res.download(outputPath, downloadName, (err) => {
    try { fs.unlinkSync(outputPath); } catch (_) {}
    if (err && !res.headersSent) {
      console.error('Download error:', err.message);
    }
  });
});

// ── START ─────────────────────────────────────────────────────────────────────
app.listen(PORT, () => {
  console.log(`CME Online Hub running on port ${PORT}`);
  console.log(`Assets: ${ASSETS_DIR}`);
  console.log(`Generator: ${GENERATOR_PATH}`);
  if (!process.env.ANTHROPIC_API_KEY) {
    console.warn('WARNING: ANTHROPIC_API_KEY is not set');
  }
});
