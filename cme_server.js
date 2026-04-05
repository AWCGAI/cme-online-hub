/**
 * CME Online Hub — Express Server
 * Collaborate Middle East PowerPoint Generator — Web Deployment
 *
 * POST /generate   → Accepts free-text brief, returns .pptx download
 * GET  /health     → Health check for Railway
 */

const express = require('express');
const path    = require('path');
const fs      = require('fs');
const { execFile } = require('child_process');
const { v4: uuidv4 } = require('uuid');
const Anthropic = require('@anthropic-ai/sdk');

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

// ── HEALTH CHECK ──────────────────────────────────────────────────────────────
app.get('/health', (req, res) => {
  res.json({ status: 'ok', timestamp: new Date().toISOString() });
});

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
bar_chart: { type, section, headline, categories: [], series: [{ name, values: [] }] }
line_chart: { type, section, headline, categories: [], series: [{ name, values: [] }] }
pie_chart: { type, section, headline, slices: [{ label, value }] }
doughnut_chart: { type, section, headline, slices: [{ label, value }] }
horiz_bar_chart: { type, section, headline, categories: [], series: [{ name, values: [] }] }
combo_chart: { type, section, headline, categories: [], bars: [{ name, values: [] }], lines: [{ name, values: [] }] }
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

## CASE STUDIES & PROJECTS

### Year of the Camel (2024) — Madinah, KSA
Saudi Arabia designated 2024 as the Year of the Camel. For Madinah, we developed a three-month interactive museum exhibit housed at the Madinah Arts Centre. Our scope covered narrative writing, design, interactive content creation, immersive storytelling and a spectacular 4D camel riding simulation. The live experience was brought to life by poets, artists and highly trained local guides.

### Madinah Climate Experience
An immersive touring visitor attraction that tells a global-scale story of the earth's climate system at risk from human inaction. Proposed at the scale of a global expo pavilion, it transports visitors into the climate system to witness its power, wonder and magnitude across the world's atmosphere and biomes.

### Al-Qaswa Museum — Celebrating Prophet Muhammad's Journey into Madinah
A permanent exhibit to enrich and expand Madinah's cultural value beyond the Year of the Camel and far into the future. This engaging, immersive exhibit takes visitors on a journey from the camel's prehistoric origins to the present day, emphasising its contribution to human civilisation and the challenges camels face today.

### Arabian Cultural Village — A Cultural Centre for Experimental Storytelling
A hub for Madinah that tells the Quba Story and educates visitors about the City Living Museum. An immersive storytelling experience to mark the start of the City Living Museum visitor journey. Reimagined as an experimental venue to test future CLM content, with detailed analysis of engagement both qualitative and quantitative informing future planning.

### CLM Visitor Centre — City Living Museum, Quba Avenue, Madinah
Creating a master hub connecting all CLM locations, experiences and commerce; providing a preview to the wonders that await around the city. A multifunctional exhibition, event and information space; a venue for history, culture and a home for live activations. The launching point for Madinah visitor exploration and business partner development.

### CLM Modular Architecture — Simplicity and Flexibility for Experiential Architecture
Creation of a unique experience architecture language for Madinah. For speed, scalability and quality control, we implemented a programme of offsite manufacture of modular units that can be arranged, rearranged and repurposed on other sites. The aim is to simplify thought processes around structure and architectural language and provide a dynamic and flexible platform to grow with demand.

### Madinah Hospitality Offer — A Cultural Hospitality Offer Unique to a City
Developing a Madinah must-have hospitality offer and product suite. Launching at CLM cafes and expanding to country, region and world. A self-funding long-term plan with economic benefit to Madinah at its core. Product development, packaging, advertising, social engagement and hospitality space design.

### Annual Pavilion Programme — A Meditation on the Future of Mosques
A new concept that focuses on the future of architecture, mosques and urbanism. It delivers international exposure for Madinah, public engagement, a potential conference subject and a lasting legacy. An annual programme to bring the Islamic world's best architects and thinkers to Madinah to create space for debate and imagination.

### Speedweek Riyadh 2024
Fusing the rich tapestry of Saudi culture with incredible cars and cutting-edge technology. An event merging the exhilaration of speed with the depths of emotion, crafting an unforgettable sensory journey. The experience synchronised heartbeats with the roar of engines, projected onto a monumental screen.

### Smart Madinah Visitor Experience
To showcase the ambitious plans of Madinah to become a leading Smart City, we developed an adaptable, digital-first space to bring stakeholders, partners and the general public together. The geometry and fluidity of the dynamic ribbon linking multiple spaces and experiences was generated via the motion capture of a leading Saudi calligraffiti artist.

### AUDI — Capturing Tomorrow, Saudi Arabia
Brief: conceptualise and execute a deeply localised production shoot showcasing the Audi brand in Saudi Arabia. Saudi Arabia is a Kingdom that embraces its cultural past while progressing towards a brighter tomorrow — the same sentiment encapsulates Audi's view of past and future. The shoot captured shared values around sustainability, design, digitalisation and performance.

### JD x Nike Retail Conference
Challenge: reinvigorate flagging Nike sales across the JD Sport retail estate. Creative idea: "Win Together" — a bespoke JD x Nike brand experience in celebratory party style, to showcase the brand's innovation and design to assembled JD retail leaders. External projections on a Manchester landmark, video walls, celebrity DJs, custom product spaces and interactive elements.

### Qiddiya — The Place to Play (Concept)
Bringing Qiddiya's thrilling entertainment to football fans through football-inspired games that capture the city's excitement. Featuring a power kick challenge, football mini-footgolf, a 360-degree selfie experience at the peak of Qiddiya's iconic Falcon rollercoaster. A taste of what's to come when the full Qiddiya entertainment district opens.

---

## UNIQUE TOOLS WE EMPLOY

### HADI — AI Cultural Guide for KSA & GCC
HADI is an AI-powered personal guide that maximises engagement for visitors to the KSA and GCC. A trip to Madinah ranks among the highlights of life for people who arrive from all around the Islamic world. HADI is a guide that understands a visitor's interests, needs, time and budgetary restrictions, and gets to know them better through a series of meaningful and relevant dialogues.

HADI features:
- Personalised AMP story discovery before arrival
- Chat-based contextual guidance on arrival
- Choice of guide persona (age and gender customisable for maximum relatability)
- Stories on Markets, Agriculture, Companions Stories, Water & Wells, Quba Mosque, and more
- Contextual familiarity: your guide speaks to you with personalised understanding

### Emotional Digital
Screenless digital experiences that create profound emotional connections. One example: an interactive carpet installation where the pattern — a map of the city — awakens at a visitor's touch, with lights rippling across it down avenues and dissipating. A sense of being somewhere we've always been going.

### Live Digital — Irresistible Theatre
Live experiences offering extraordinary possibilities for digital interaction. Blending screens, devices and physical experiences to create irresistible theatre. When spectacular experiences are combined with breathtaking sound, the best lighting and special effects, audiences are emotionally affected. Irresistible theatre drives sharing which drives memory.

---

## LEADERSHIP TEAM

**Ben McMahon** — Global CEO & Founder — ben@collaborateglobal.com
**Nick Walsh** — Client Experience Director — nick@collaborateglobal.com
**Andrew Walker** — Chief Creative Officer — andrew@collaborateglobal.com
**Natassha Evans** — Head of Client Experience — natassha@collaborateglobal.com
**Ewan Ferrier** — Creative Director — ewan@collaborateglobal.com
**Hamad Tariq Mahmoud** — Saudi Cultural Lead — hamid@collaborateglobal.com
**Hamsa Amjed** — Global Consultant — hamsa@collaborateglobal.com
**Ross Oxenham** — Chief Expansion Officer — ross@collaborateglobal.com
**Abdulrahman Alrashidi** — Agency Manager — abdulrahman@collaborateglobal.com

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

    // Read the generator source
    const generatorSource = fs.readFileSync(GENERATOR_PATH, 'utf8');

    // Find the cut point — everything before the BRIEF comment
    const BRIEF_MARKER = '// BRIEF — EDIT THIS SECTION TO PRODUCE A NEW DECK';
    const cutIndex = generatorSource.indexOf(BRIEF_MARKER);
    if (cutIndex === -1) {
      return reject(new Error('Could not find BRIEF marker in cme_generator.js'));
    }

    let engineCode = generatorSource.slice(0, cutIndex);

    // Override the hardcoded paths
    engineCode = engineCode
      .replace(
        'const ASSETS_DIR = "/home/claude/cme_assets"',
        `const ASSETS_DIR = ${JSON.stringify(ASSETS_DIR)}`
      )
      .replace(
        'const OUTPUT_PATH = "/home/claude/cg_output.pptx"',
        `const OUTPUT_PATH = ${JSON.stringify(outputPath)}`
      );

    // Build the runner: engine + new BRIEF + buildDeck call
    const runner = engineCode +
      `\n// BRIEF — EDIT THIS SECTION TO PRODUCE A NEW DECK\n` +
      `const BRIEF = ${briefText};\n\n` +
      `buildDeck(BRIEF);\n`;

    fs.writeFileSync(runnerPath, runner, 'utf8');

    execFile('node', [runnerPath], {
      cwd: __dirname,
      timeout: 120000,
      maxBuffer: 10 * 1024 * 1024
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

// ── GENERATE ENDPOINT ─────────────────────────────────────────────────────────
app.post('/generate', async (req, res) => {
  const { content } = req.body;

  if (!content || content.trim().length < 10) {
    return res.status(400).json({ error: 'Please provide content to convert into a presentation.' });
  }

  if (!process.env.ANTHROPIC_API_KEY) {
    return res.status(500).json({ error: 'ANTHROPIC_API_KEY environment variable is not set.' });
  }

  console.log(`[${new Date().toISOString()}] Generating brief for content (${content.length} chars)`);

  let briefText, outputPath;

  try {
    // Step 1: Generate BRIEF via Haiku
    const result = await generateBrief(content);
    briefText = result.briefText;
    console.log(`Brief generated: ${result.brief.slides.length} slides`);
  } catch (err) {
    console.error('Brief generation failed:', err.message);
    return res.status(500).json({ error: `Brief generation failed: ${err.message}` });
  }

  try {
    // Step 2: Build the .pptx
    outputPath = await buildPptx(briefText);
    console.log(`PPTX built: ${outputPath}`);
  } catch (err) {
    console.error('PPTX build failed:', err.message);
    return res.status(500).json({ error: `PPTX build failed: ${err.message}` });
  }

  // Step 3: Stream the file then clean up
  res.download(outputPath, 'collaborate-me-presentation.pptx', (err) => {
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
