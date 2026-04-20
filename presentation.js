"use strict";
const pptxgen = require("pptxgenjs");

// ── Palette ───────────────────────────────────────────────────────────────────
const DARK    = "1E2761";   // navy
const ACCENT  = "CADCFC";   // ice blue
const WHITE   = "FFFFFF";
const LIGHT   = "F4F7FF";   // slide background (light slides)
const CARD    = "EAEFFF";   // inner card fill
const SHADOW_DARK = "141A45";
const TEAL    = "1A6B4A";
const BERRY   = "6B2E5A";
const PURPLE  = "4A2E8E";
const OLIVE   = "4A6B2E";
const RUST    = "8E3A2E";

// ── Helpers ───────────────────────────────────────────────────────────────────
const mkShadow = () => ({ type: "outer", blur: 5, offset: 2, angle: 135, color: "000000", opacity: 0.13 });

function leftBar(s, color) {
  color = color || DARK;
  s.addShape("rect", { x: 0, y: 0, w: 0.2, h: 5.625, fill: { color } });
}

function slideTitle(s, text) {
  s.addText(text, {
    x: 0.42, y: 0.22, w: 9.16, h: 0.62,
    fontSize: 34, bold: true, color: DARK, fontFace: "Calibri", margin: 0,
  });
}

// Two-column card — returns nothing, mutates slide
function twoCard(s, left, right) {
  // left card
  s.addShape("rect", { x: 0.42, y: 1.05, w: 4.35, h: 4.1, fill: { color: DARK }, shadow: mkShadow() });
  s.addShape("rect", { x: 0.42, y: 1.05, w: 4.35, h: 0.52, fill: { color: left.accent || ACCENT } });
  s.addText(left.title, { x: 0.52, y: 1.05, w: 4.15, h: 0.52, fontSize: 16, bold: true, color: DARK, fontFace: "Calibri", valign: "middle", margin: 0 });
  s.addText(left.body, { x: 0.6, y: 1.7, w: 4.05, h: 3.3, fontSize: 13, color: WHITE, fontFace: "Calibri", valign: "top" });

  // right card
  s.addShape("rect", { x: 5.22, y: 1.05, w: 4.35, h: 4.1, fill: { color: DARK }, shadow: mkShadow() });
  s.addShape("rect", { x: 5.22, y: 1.05, w: 4.35, h: 0.52, fill: { color: right.accent || ACCENT } });
  s.addText(right.title, { x: 5.32, y: 1.05, w: 4.15, h: 0.52, fontSize: 16, bold: true, color: DARK, fontFace: "Calibri", valign: "middle", margin: 0 });
  s.addText(right.body, { x: 5.32, y: 1.7, w: 4.05, h: 3.3, fontSize: 13, color: WHITE, fontFace: "Calibri", valign: "top" });
}

// Section / category card (full dark slide)
function sectionSlide(pres, category, color, count, description, skills) {
  const s = pres.addSlide();
  s.background = { color };
  s.addShape("rect", { x: 0, y: 0, w: 0.45, h: 5.625, fill: { color: ACCENT } });
  s.addText(category.toUpperCase(), {
    x: 0.65, y: 0.55, w: 8.9, h: 0.5,
    fontSize: 12, bold: true, color: ACCENT, fontFace: "Calibri", charSpacing: 4, margin: 0,
  });
  s.addText(category, {
    x: 0.65, y: 1.0, w: 8.9, h: 1.35,
    fontSize: 52, bold: true, color: WHITE, fontFace: "Calibri", margin: 0,
  });
  // Count badge
  s.addShape("rect", { x: 0.65, y: 2.48, w: 1.3, h: 0.46, fill: { color: ACCENT } });
  s.addText(`${count} Skills`, { x: 0.65, y: 2.48, w: 1.3, h: 0.46, fontSize: 13, bold: true, color: DARK, align: "center", valign: "middle", fontFace: "Calibri", margin: 0 });

  s.addText(description, {
    x: 0.65, y: 3.08, w: 5.5, h: 0.85,
    fontSize: 16, color: ACCENT, fontFace: "Calibri", margin: 0,
  });
  // Skills list on right
  const skillStr = skills.map(sk => "  " + sk).join("\n");
  s.addShape("rect", { x: 6.5, y: 1.5, w: 3.1, h: 3.7, fill: { color: "000000", transparency: 60 } });
  skills.forEach((sk, i) => {
    s.addText(sk, {
      x: 6.7, y: 1.65 + i * 0.55, w: 2.85, h: 0.42,
      fontSize: 13, color: WHITE, fontFace: "Calibri", margin: 0,
      bullet: false,
    });
  });
}

// ─────────────────────────────────────────────────────────────────────────────
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title  = "Claude Agent Skills";

// ── Slide 1 · Title ───────────────────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: DARK };
  // Bottom band
  s.addShape("rect", { x: 0, y: 3.9, w: 10, h: 1.725, fill: { color: SHADOW_DARK } });
  // Left accent
  s.addShape("rect", { x: 0, y: 0, w: 0.55, h: 5.625, fill: { color: ACCENT } });
  // Eyebrow
  s.addText("ANTHROPIC  ·  AGENT SKILLS", {
    x: 0.8, y: 0.42, w: 8.8, h: 0.38,
    fontSize: 11, bold: true, color: ACCENT, fontFace: "Calibri", charSpacing: 5, margin: 0,
  });
  // Main title
  s.addText("Claude Agent Skills", {
    x: 0.8, y: 0.9, w: 8.8, h: 1.6,
    fontSize: 58, bold: true, color: WHITE, fontFace: "Calibri", margin: 0,
  });
  // Subtitle
  s.addText("A Production-Grade Toolkit for Extending AI Capabilities", {
    x: 0.8, y: 2.58, w: 8.8, h: 0.65,
    fontSize: 21, color: ACCENT, fontFace: "Calibri", margin: 0,
  });
  // Stats row
  [["17", "Skills"], ["5", "Categories"], ["Production", "Grade"]].forEach(([n, l], i) => {
    const x = 0.8 + i * 3.0;
    s.addText(n,  { x, y: 4.08, w: 2.8, h: 0.6,  fontSize: 30, bold: true, color: WHITE,  fontFace: "Calibri", margin: 0 });
    s.addText(l,  { x, y: 4.65, w: 2.8, h: 0.38, fontSize: 13, color: ACCENT, fontFace: "Calibri", margin: 0 });
  });
}

// ── Slide 2 · What Are Skills? ────────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: LIGHT };
  leftBar(s);
  slideTitle(s, "What Are Agent Skills?");
  // Divider
  s.addShape("rect", { x: 0.42, y: 0.9, w: 9.16, h: 0.05, fill: { color: ACCENT } });

  // Left feature block
  s.addShape("rect", { x: 0.42, y: 1.08, w: 4.55, h: 4.1, fill: { color: DARK }, shadow: mkShadow() });
  s.addText([
    { text: "Skills are structured instruction sets", options: { bold: true, breakLine: true } },
    { text: "that extend Claude's capabilities for specialized, production-grade tasks.", options: { breakLine: true } },
    { text: " ", options: { breakLine: true } },
    { text: "Each skill bundles:", options: { bold: true, breakLine: true } },
    { text: "Intent & workflow guidance", options: { bullet: true, breakLine: true } },
    { text: "Curated tool recommendations", options: { bullet: true, breakLine: true } },
    { text: "Pitfall warnings & best practices", options: { bullet: true, breakLine: true } },
    { text: "Supporting scripts & templates", options: { bullet: true, breakLine: true } },
    { text: " ", options: { breakLine: true } },
    { text: "Skills live in SKILL.md files with YAML front-matter that lets Claude trigger them automatically from natural language.", options: {} },
  ], { x: 0.58, y: 1.22, w: 4.25, h: 3.82, fontSize: 14, color: WHITE, fontFace: "Calibri", valign: "top" });

  // Right — three fact cards
  [
    { tag: "Trigger-Based",  desc: "Claude activates a skill automatically when your request matches its description — no slash command required." },
    { tag: "Composable",     desc: "Skills layer on top of each other. Use PPTX + Brand Guidelines together for branded decks." },
    { tag: "Extensible",     desc: "The Skill Creator skill lets you build, benchmark, and refine entirely new skills." },
  ].forEach(({ tag, desc }, i) => {
    const y = 1.08 + i * 1.4;
    s.addShape("rect", { x: 5.3, y, w: 4.28, h: 1.24, fill: { color: WHITE }, shadow: mkShadow() });
    s.addShape("rect", { x: 5.3, y, w: 0.14, h: 1.24, fill: { color: ACCENT } });
    s.addText(tag,  { x: 5.56, y: y + 0.1,  w: 3.9, h: 0.36, fontSize: 17, bold: true, color: DARK,    fontFace: "Calibri", margin: 0 });
    s.addText(desc, { x: 5.56, y: y + 0.5,  w: 3.9, h: 0.64, fontSize: 12, color: "444466", fontFace: "Calibri", margin: 0 });
  });
}

// ── Slide 3 · Skill Landscape ─────────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: DARK };
  s.addText("The Complete Skill Landscape", {
    x: 0.5, y: 0.18, w: 9, h: 0.58, fontSize: 32, bold: true, color: WHITE, fontFace: "Calibri", margin: 0,
  });
  s.addText("17 production-grade skills across 5 categories", {
    x: 0.5, y: 0.75, w: 9, h: 0.32, fontSize: 15, color: ACCENT, fontFace: "Calibri", margin: 0,
  });

  const cats = [
    { name: "Creative\n& Design",  color: "2E4A8E", accent: "9BB5F5", skills: ["Algorithmic Art", "Canvas Design", "Frontend Design", "Theme Factory", "Slack GIF Creator"] },
    { name: "Development\n& Tech",  color: "1A5C42", accent: "7EC8A0", skills: ["Claude API", "MCP Builder", "Skill Creator", "Web App Testing"] },
    { name: "Enterprise\n& Comms",  color: "7A3020", accent: "F5A87E", skills: ["Brand Guidelines", "Internal Comms"] },
    { name: "Document\nSuite",      color: "4A2C82", accent: "B8A0F0", skills: ["DOCX", "PDF", "PPTX", "XLSX"] },
    { name: "Specialized",          color: "3A5A1A", accent: "A8CC7A", skills: ["Doc Coauthoring", "Web Artifacts Builder"] },
  ];

  cats.forEach(({ name, color, accent, skills }, i) => {
    const x = 0.22 + i * 1.93;
    const w = 1.76;
    s.addShape("rect", { x, y: 1.22, w, h: 4.0, fill: { color }, shadow: mkShadow() });
    s.addShape("rect", { x, y: 1.22, w, h: 0.6, fill: { color: accent } });
    s.addText(name, { x, y: 1.22, w, h: 0.6, fontSize: 11, bold: true, color: "1A1A2E", align: "center", valign: "middle", fontFace: "Calibri", margin: 2 });
    skills.forEach((sk, j) => {
      s.addShape("rect", { x: x + 0.08, y: 1.96 + j * 0.58, w: w - 0.16, h: 0.46, fill: { color: "000000", transparency: 50 } });
      s.addText(sk, { x: x + 0.12, y: 1.97 + j * 0.58, w: w - 0.22, h: 0.44, fontSize: 10.5, color: WHITE, fontFace: "Calibri", valign: "middle", margin: 0 });
    });
  });
}

// ── Slide 4 · Section: Creative & Design ─────────────────────────────────────
sectionSlide(pres, "Creative & Design", "2E4A8E", 5,
  "Tools for generating visual art, web interfaces,\ndesigns, themes, and animations.",
  ["Algorithmic Art", "Canvas Design", "Frontend Design", "Theme Factory", "Slack GIF Creator"]);

// ── Slide 5 · Algorithmic Art + Canvas Design ─────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: LIGHT };
  leftBar(s, "2E4A8E");
  slideTitle(s, "Algorithmic Art  ·  Canvas Design");
  s.addShape("rect", { x: 0.42, y: 0.9, w: 9.16, h: 0.05, fill: { color: "9BB5F5" } });

  twoCard(s,
    {
      title: "Algorithmic Art",
      accent: "9BB5F5",
      body: [
        { text: "Creates generative art using p5.js with seeded randomness — outputs are reproducible yet feel alive.", options: { breakLine: true } },
        { text: " ", options: { breakLine: true } },
        { text: "Two-phase process:", options: { bold: true, breakLine: true } },
        { text: "1. Artistic philosophy markdown", options: { breakLine: true } },
        { text: "2. Interactive HTML artifact", options: { breakLine: true } },
        { text: " ", options: { breakLine: true } },
        { text: "Focus: computational beauty, parametric variation, emergent behavior, and controlled chaos.", options: {} },
      ].map(b => ({ ...b })),
    },
    {
      title: "Canvas Design",
      accent: "9BB5F5",
      body: [
        { text: "Produces static visual designs as PDFs or PNGs — think coffee-table book pages and gallery prints.", options: { breakLine: true } },
        { text: " ", options: { breakLine: true } },
        { text: "Two-phase process:", options: { bold: true, breakLine: true } },
        { text: "1. Design philosophy markdown", options: { breakLine: true } },
        { text: "2. Visual PDF or PNG output", options: { breakLine: true } },
        { text: " ", options: { breakLine: true } },
        { text: "Core principle: minimalist text, maximum visual impact through font, color, form, and composition.", options: {} },
      ].map(b => ({ ...b })),
    }
  );
}

// ── Slide 6 · Frontend Design + Theme Factory ─────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: LIGHT };
  leftBar(s, "2E4A8E");
  slideTitle(s, "Frontend Design  ·  Theme Factory");
  s.addShape("rect", { x: 0.42, y: 0.9, w: 9.16, h: 0.05, fill: { color: "9BB5F5" } });

  twoCard(s,
    {
      title: "Frontend Design",
      accent: "9BB5F5",
      body: [
        { text: "Produces distinctive, production-grade web interfaces using HTML, CSS, JavaScript, React, or Vue.", options: { breakLine: true } },
        { text: " ", options: { breakLine: true } },
        { text: "Key design rules:", options: { bold: true, breakLine: true } },
        { text: "Bold aesthetic direction — avoid generic AI slop", options: { bullet: true, breakLine: true } },
        { text: "Distinctive typography (not default Inter)", options: { bullet: true, breakLine: true } },
        { text: "Intentional motion & spatial composition", options: { bullet: true, breakLine: true } },
        { text: "Exceptional visual quality with working code", options: { bullet: true } },
      ].map(b => ({ ...b })),
    },
    {
      title: "Theme Factory",
      accent: "9BB5F5",
      body: [
        { text: "Applies professional themes to any artifact — slides, docs, reports, or HTML pages.", options: { breakLine: true } },
        { text: " ", options: { breakLine: true } },
        { text: "10 curated built-in themes:", options: { bold: true, breakLine: true } },
        { text: "Ocean Depths, Sunset Boulevard, Arctic Frost,", options: { breakLine: true } },
        { text: "and 7 more — each with color palettes + font pairings.", options: { breakLine: true } },
        { text: " ", options: { breakLine: true } },
        { text: "Custom themes can be generated on-the-fly from any brand or creative brief.", options: {} },
      ].map(b => ({ ...b })),
    }
  );
}

// ── Slide 7 · Slack GIF Creator ───────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: LIGHT };
  leftBar(s, "2E4A8E");
  slideTitle(s, "Slack GIF Creator");
  s.addShape("rect", { x: 0.42, y: 0.9, w: 9.16, h: 0.05, fill: { color: "9BB5F5" } });

  // Feature block left
  s.addShape("rect", { x: 0.42, y: 1.08, w: 4.55, h: 4.1, fill: { color: "2E4A8E" }, shadow: mkShadow() });
  s.addText([
    { text: "Generates animated GIFs optimized for Slack.", options: { bold: true, breakLine: true } },
    { text: " ", options: { breakLine: true } },
    { text: "Two output sizes:", options: { bold: true, breakLine: true } },
    { text: "Emoji GIFs — 128 × 128 px", options: { bullet: true, breakLine: true } },
    { text: "Message GIFs — 480 × 480 px", options: { bullet: true, breakLine: true } },
    { text: " ", options: { breakLine: true } },
    { text: "Built-in animation concepts:", options: { bold: true, breakLine: true } },
    { text: "shake, pulse, bounce, spin, fade, slide, zoom, explode", options: { breakLine: true } },
    { text: " ", options: { breakLine: true } },
    { text: "Library includes easing functions, frame helpers, and validators — all optimized for file-size limits.", options: {} },
  ], { x: 0.6, y: 1.22, w: 4.22, h: 3.82, fontSize: 14, color: WHITE, fontFace: "Calibri", valign: "top" });

  // Right — Animation concepts grid
  const concepts = ["shake", "pulse", "bounce", "spin", "fade", "slide", "zoom", "explode"];
  s.addText("Animation Concepts", { x: 5.3, y: 1.08, w: 4.28, h: 0.44, fontSize: 16, bold: true, color: DARK, fontFace: "Calibri", margin: 0 });
  concepts.forEach((c, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = 5.3 + col * 2.15;
    const y = 1.62 + row * 0.85;
    s.addShape("rect", { x, y, w: 2.0, h: 0.68, fill: { color: "2E4A8E" }, shadow: mkShadow() });
    s.addText(c, { x, y, w: 2.0, h: 0.68, fontSize: 15, bold: true, color: WHITE, align: "center", valign: "middle", fontFace: "Calibri", margin: 0 });
  });

  s.addText("PIL-based frame generation · Auto file-size optimization · Built-in validators", {
    x: 5.3, y: 5.1, w: 4.28, h: 0.38, fontSize: 11, color: "666688", fontFace: "Calibri", margin: 0,
  });
}

// ── Slide 8 · Section: Development & Technical ───────────────────────────────
sectionSlide(pres, "Development & Technical", "1A5C42", 4,
  "Guides and workflows for building LLM apps,\nMCP servers, skills, and automated tests.",
  ["Claude API", "MCP Builder", "Skill Creator", "Web App Testing"]);

// ── Slide 9 · Claude API ──────────────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: LIGHT };
  leftBar(s, "1A5C42");
  slideTitle(s, "Claude API Skill");
  s.addShape("rect", { x: 0.42, y: 0.9, w: 9.16, h: 0.05, fill: { color: "7EC8A0" } });

  // Main description block
  s.addShape("rect", { x: 0.42, y: 1.08, w: 9.16, h: 1.0, fill: { color: "1A5C42" }, shadow: mkShadow() });
  s.addText("Comprehensive guide for building LLM-powered applications — covers every SDK surface across 7 languages.", {
    x: 0.6, y: 1.08, w: 8.8, h: 1.0, fontSize: 15, color: WHITE, fontFace: "Calibri", valign: "middle",
  });

  // Four feature cards
  const apiCards = [
    { title: "7 SDKs",       desc: "Python, TypeScript, Java, Go, Ruby, C#, PHP — with language-specific guidance for each." },
    { title: "All Surfaces", desc: "Single calls, tool use, streaming, agents, Managed Agents, batching, structured output." },
    { title: "Key Features", desc: "Prompt caching, extended thinking, compaction, citations, Files API, model selection." },
    { title: "Current Models", desc: "Opus 4.7, Sonnet 4.6, Haiku 4.5 — with decision trees for choosing the right one." },
  ];
  apiCards.forEach(({ title, desc }, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = 0.42 + col * 4.65, y = 2.28 + row * 1.55;
    s.addShape("rect", { x, y, w: 4.48, h: 1.38, fill: { color: WHITE }, shadow: mkShadow() });
    s.addShape("rect", { x, y, w: 0.12, h: 1.38, fill: { color: "7EC8A0" } });
    s.addText(title, { x: x + 0.22, y: y + 0.1, w: 4.1, h: 0.36, fontSize: 16, bold: true, color: "1A5C42", fontFace: "Calibri", margin: 0 });
    s.addText(desc,  { x: x + 0.22, y: y + 0.5, w: 4.1, h: 0.78, fontSize: 12.5, color: "333355", fontFace: "Calibri", margin: 0 });
  });
}

// ── Slide 10 · MCP Builder + Skill Creator ────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: LIGHT };
  leftBar(s, "1A5C42");
  slideTitle(s, "MCP Builder  ·  Skill Creator");
  s.addShape("rect", { x: 0.42, y: 0.9, w: 9.16, h: 0.05, fill: { color: "7EC8A0" } });

  twoCard(s,
    {
      title: "MCP Builder",
      accent: "7EC8A0",
      body: [
        { text: "Guides you through building Model Context Protocol servers — the standard for connecting AI to external tools and data.", options: { breakLine: true } },
        { text: " ", options: { breakLine: true } },
        { text: "Four-phase workflow:", options: { bold: true, breakLine: true } },
        { text: "Research → Implementation → Review → Evaluate", options: { breakLine: true } },
        { text: " ", options: { breakLine: true } },
        { text: "Recommended stack:", options: { bold: true, breakLine: true } },
        { text: "TypeScript + HTTP transport for remote servers", options: { bullet: true, breakLine: true } },
        { text: "stdio transport for local integrations", options: { bullet: true } },
      ].map(b => ({ ...b })),
    },
    {
      title: "Skill Creator",
      accent: "7EC8A0",
      body: [
        { text: "A meta-skill for building, benchmarking, and refining other skills — the toolkit for extending the toolkit.", options: { breakLine: true } },
        { text: " ", options: { breakLine: true } },
        { text: "Workflow:", options: { bold: true, breakLine: true } },
        { text: "Capture intent → Write draft → Run tests → Evaluate → Iterate", options: { breakLine: true } },
        { text: " ", options: { breakLine: true } },
        { text: "Includes benchmark viewer with qualitative and quantitative metrics, and description optimization to improve trigger accuracy.", options: {} },
      ].map(b => ({ ...b })),
    }
  );
}

// ── Slide 11 · Web App Testing ────────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: LIGHT };
  leftBar(s, "1A5C42");
  slideTitle(s, "Web App Testing");
  s.addShape("rect", { x: 0.42, y: 0.9, w: 9.16, h: 0.05, fill: { color: "7EC8A0" } });

  s.addShape("rect", { x: 0.42, y: 1.08, w: 4.55, h: 4.1, fill: { color: "1A5C42" }, shadow: mkShadow() });
  s.addText([
    { text: "Playwright-based testing for local web apps — both static HTML and dynamic single-page apps.", options: { breakLine: true } },
    { text: " ", options: { breakLine: true } },
    { text: "Reconnaissance-then-action pattern:", options: { bold: true, breakLine: true } },
    { text: "Inspect DOM → Identify selectors → Execute tests", options: { breakLine: true } },
    { text: " ", options: { breakLine: true } },
    { text: "Captures:", options: { bold: true, breakLine: true } },
    { text: "Screenshots at each step", options: { bullet: true, breakLine: true } },
    { text: "Browser console logs", options: { bullet: true, breakLine: true } },
    { text: "Full DOM snapshots", options: { bullet: true } },
  ], { x: 0.6, y: 1.22, w: 4.22, h: 3.82, fontSize: 14, color: WHITE, fontFace: "Calibri", valign: "top" });

  const steps = [
    { n: "1", label: "Start Server",    desc: "Helper scripts manage server lifecycle automatically" },
    { n: "2", label: "Inspect DOM",     desc: "Identify selectors, routes, and interactive elements" },
    { n: "3", label: "Execute Tests",   desc: "Run Playwright actions across flows and edge cases" },
    { n: "4", label: "Capture Evidence",desc: "Screenshots + console logs saved per test step" },
  ];
  steps.forEach(({ n, label, desc }, i) => {
    const y = 1.08 + i * 1.05;
    s.addShape("rect", { x: 5.3, y, w: 4.28, h: 0.9, fill: { color: WHITE }, shadow: mkShadow() });
    s.addShape("rect", { x: 5.3, y, w: 0.55, h: 0.9, fill: { color: "1A5C42" } });
    s.addText(n, { x: 5.3, y, w: 0.55, h: 0.9, fontSize: 20, bold: true, color: WHITE, align: "center", valign: "middle", fontFace: "Calibri", margin: 0 });
    s.addText(label, { x: 5.98, y: y + 0.08, w: 3.5, h: 0.32, fontSize: 15, bold: true, color: "1A5C42", fontFace: "Calibri", margin: 0 });
    s.addText(desc,  { x: 5.98, y: y + 0.44, w: 3.5, h: 0.38, fontSize: 11.5, color: "444466", fontFace: "Calibri", margin: 0 });
  });
}

// ── Slide 12 · Section: Enterprise & Communication ────────────────────────────
sectionSlide(pres, "Enterprise & Communication", "7A3020", 2,
  "Brand-consistent styling and structured\nworkflows for company-wide communications.",
  ["Brand Guidelines", "Internal Comms"]);

// ── Slide 13 · Brand Guidelines + Internal Comms ─────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: LIGHT };
  leftBar(s, "7A3020");
  slideTitle(s, "Brand Guidelines  ·  Internal Communications");
  s.addShape("rect", { x: 0.42, y: 0.9, w: 9.16, h: 0.05, fill: { color: "F5A87E" } });

  twoCard(s,
    {
      title: "Brand Guidelines",
      accent: "F5A87E",
      body: [
        { text: "Applies Anthropic's official brand styling to any output — colors, typography, and accent system.", options: { breakLine: true } },
        { text: " ", options: { breakLine: true } },
        { text: "Core palette:", options: { bold: true, breakLine: true } },
        { text: "Dark #141413  ·  Light #faf9f5", options: { breakLine: true } },
        { text: "Orange #d97757  ·  Blue #6a9bcc  ·  Green #788c5d", options: { breakLine: true } },
        { text: " ", options: { breakLine: true } },
        { text: "Typography:", options: { bold: true, breakLine: true } },
        { text: "Poppins (headings)  ·  Lora (body)", options: { breakLine: true } },
        { text: "Smart font fallbacks built in", options: {} },
      ].map(b => ({ ...b })),
    },
    {
      title: "Internal Communications",
      accent: "F5A87E",
      body: [
        { text: "Structured workflows for company-wide communication — standardizes tone and format across the org.", options: { breakLine: true } },
        { text: " ", options: { breakLine: true } },
        { text: "Supported formats:", options: { bold: true, breakLine: true } },
        { text: "3P Updates (Progress / Plans / Problems)", options: { bullet: true, breakLine: true } },
        { text: "Newsletters & team updates", options: { bullet: true, breakLine: true } },
        { text: "FAQs & status reports", options: { bullet: true, breakLine: true } },
        { text: "Incident reports", options: { bullet: true } },
      ].map(b => ({ ...b })),
    }
  );
}

// ── Slide 14 · Section: Document Suite ───────────────────────────────────────
sectionSlide(pres, "Document Suite", "4A2C82", 4,
  "Create, read, and edit Word, PDF, PowerPoint,\nand Excel files with expert-level guidance.",
  ["DOCX", "PDF", "PPTX", "XLSX"]);

// ── Slide 15 · DOCX + PDF ─────────────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: LIGHT };
  leftBar(s, "4A2C82");
  slideTitle(s, "DOCX  ·  PDF");
  s.addShape("rect", { x: 0.42, y: 0.9, w: 9.16, h: 0.05, fill: { color: "B8A0F0" } });

  twoCard(s,
    {
      title: "DOCX — Word Documents",
      accent: "B8A0F0",
      body: [
        { text: "Full create / read / edit lifecycle for Word documents, using docx-js for creation and pandoc for extraction.", options: { breakLine: true } },
        { text: " ", options: { breakLine: true } },
        { text: "Advanced features:", options: { bold: true, breakLine: true } },
        { text: "Tables, headers/footers, TOC, tracked changes", options: { bullet: true, breakLine: true } },
        { text: "Comments, images, and embedded objects", options: { bullet: true, breakLine: true } },
        { text: " ", options: { breakLine: true } },
        { text: "Critical notes:", options: { bold: true, breakLine: true } },
        { text: "Defaults to A4 — specify US Letter explicitly", options: { bullet: true, breakLine: true } },
        { text: "Validation scripts auto-repair common issues", options: { bullet: true } },
      ].map(b => ({ ...b })),
    },
    {
      title: "PDF — Full Operations",
      accent: "B8A0F0",
      body: [
        { text: "Comprehensive PDF workflows using pypdf, pdfplumber, and reportlab.", options: { breakLine: true } },
        { text: " ", options: { breakLine: true } },
        { text: "Capabilities:", options: { bold: true, breakLine: true } },
        { text: "Merge, split, rotate, extract text/tables", options: { bullet: true, breakLine: true } },
        { text: "OCR, watermark, password protect", options: { bullet: true, breakLine: true } },
        { text: "Create from scratch with reportlab", options: { bullet: true, breakLine: true } },
        { text: " ", options: { breakLine: true } },
        { text: "CLI tools: pdftotext, qpdf, pdftk", options: {} },
      ].map(b => ({ ...b })),
    }
  );
}

// ── Slide 16 · PPTX + XLSX ───────────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: LIGHT };
  leftBar(s, "4A2C82");
  slideTitle(s, "PPTX  ·  XLSX");
  s.addShape("rect", { x: 0.42, y: 0.9, w: 9.16, h: 0.05, fill: { color: "B8A0F0" } });

  twoCard(s,
    {
      title: "PPTX — PowerPoint",
      accent: "B8A0F0",
      body: [
        { text: "Create and edit PowerPoint presentations with bold design guidance and mandatory visual QA.", options: { breakLine: true } },
        { text: " ", options: { breakLine: true } },
        { text: "Key principles:", options: { bold: true, breakLine: true } },
        { text: "Vary layouts — never repeat the same slide", options: { bullet: true, breakLine: true } },
        { text: "Every slide needs a visual element", options: { bullet: true, breakLine: true } },
        { text: "Assume problems — use subagents for QA", options: { bullet: true, breakLine: true } },
        { text: " ", options: { breakLine: true } },
        { text: "Tools: pptxgenjs · markitdown · LibreOffice", options: {} },
      ].map(b => ({ ...b })),
    },
    {
      title: "XLSX — Spreadsheets",
      accent: "B8A0F0",
      body: [
        { text: "Create and edit Excel files using pandas for analysis and openpyxl for formulas and formatting.", options: { breakLine: true } },
        { text: " ", options: { breakLine: true } },
        { text: "Critical rules:", options: { bold: true, breakLine: true } },
        { text: "Always use formulas — never hardcode values", options: { bullet: true, breakLine: true } },
        { text: "Zero formula errors before delivery", options: { bullet: true, breakLine: true } },
        { text: "Run recalc.py after formula creation", options: { bullet: true, breakLine: true } },
        { text: " ", options: { breakLine: true } },
        { text: "Color standard: blue=inputs, black=formulas, green=internal, red=external", options: {} },
      ].map(b => ({ ...b })),
    }
  );
}

// ── Slide 17 · Section: Specialized ──────────────────────────────────────────
sectionSlide(pres, "Specialized Skills", "3A5A1A", 2,
  "Collaborative document creation and complex\nReact/TypeScript artifacts for Claude.ai.",
  ["Doc Coauthoring", "Web Artifacts Builder"]);

// ── Slide 18 · Doc Coauthoring + Web Artifacts Builder ───────────────────────
{
  const s = pres.addSlide();
  s.background = { color: LIGHT };
  leftBar(s, "3A5A1A");
  slideTitle(s, "Doc Coauthoring  ·  Web Artifacts Builder");
  s.addShape("rect", { x: 0.42, y: 0.9, w: 9.16, h: 0.05, fill: { color: "A8CC7A" } });

  twoCard(s,
    {
      title: "Doc Coauthoring",
      accent: "A8CC7A",
      body: [
        { text: "Structured three-stage workflow for collaborative document creation — catches blind spots by testing with a fresh Claude instance.", options: { breakLine: true } },
        { text: " ", options: { breakLine: true } },
        { text: "Three stages:", options: { bold: true, breakLine: true } },
        { text: "1. Context Gathering", options: { bullet: true, breakLine: true } },
        { text: "2. Refinement & Structure", options: { bullet: true, breakLine: true } },
        { text: "3. Reader Testing (sub-agent)", options: { bullet: true, breakLine: true } },
        { text: " ", options: { breakLine: true } },
        { text: "Progressive disclosure approach — clarity for readers, not just authors.", options: {} },
      ].map(b => ({ ...b })),
    },
    {
      title: "Web Artifacts Builder",
      accent: "A8CC7A",
      body: [
        { text: "Builds complex React/TypeScript artifacts for Claude.ai — bundled to a single self-contained HTML file.", options: { breakLine: true } },
        { text: " ", options: { breakLine: true } },
        { text: "Stack:", options: { bold: true, breakLine: true } },
        { text: "React 18 + TypeScript + Vite + Parcel", options: { bullet: true, breakLine: true } },
        { text: "Tailwind CSS + shadcn/ui", options: { bullet: true, breakLine: true } },
        { text: " ", options: { breakLine: true } },
        { text: "Design rules:", options: { bold: true, breakLine: true } },
        { text: "No excessive centering, purple gradients, or Inter — avoid AI slop aesthetics.", options: {} },
      ].map(b => ({ ...b })),
    }
  );
}

// ── Slide 19 · Key Design Themes ─────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: LIGHT };
  leftBar(s, DARK);
  slideTitle(s, "Key Design Themes Across All Skills");
  s.addShape("rect", { x: 0.42, y: 0.9, w: 9.16, h: 0.05, fill: { color: ACCENT } });

  const themes = [
    { title: "Two-Phase Approach",    desc: "Creative skills output a philosophy doc first, then the artifact — intent guides execution." },
    { title: "Craftsmanship Standard",desc: "Every skill targets master-level output. Generic 'AI slop' aesthetics are explicitly forbidden." },
    { title: "QA as Primary Task",    desc: "PPTX requires subagent visual inspection. XLSX requires zero errors. Assume problems exist." },
    { title: "Composable by Design",  desc: "Skills layer: PPTX + Brand Guidelines = branded decks. Claude API + MCP Builder = full stack." },
    { title: "Template-Based Start",  desc: "Algorithmic Art, Canvas, and Web Artifacts all provide templates — never start from scratch." },
    { title: "Multi-Format Output",   desc: "Skills routinely produce 2+ file types (e.g., .md philosophy + .html artifact) for richer results." },
  ];

  themes.forEach(({ title, desc }, i) => {
    const col = i % 3, row = Math.floor(i / 2);
    const x = 0.42 + col * 3.1, y = 1.1 + row * 1.6;
    s.addShape("rect", { x, y, w: 2.95, h: 1.45, fill: { color: DARK }, shadow: mkShadow() });
    s.addShape("rect", { x, y, w: 2.95, h: 0.38, fill: { color: ACCENT } });
    s.addText(title, { x: x + 0.1, y, w: 2.75, h: 0.38, fontSize: 12, bold: true, color: DARK, fontFace: "Calibri", valign: "middle", margin: 0 });
    s.addText(desc,  { x: x + 0.1, y: y + 0.45, w: 2.75, h: 0.9, fontSize: 11.5, color: WHITE, fontFace: "Calibri", valign: "top", margin: 0 });
  });
}

// ── Slide 20 · Conclusion ─────────────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: DARK };
  s.addShape("rect", { x: 0, y: 0, w: 0.55, h: 5.625, fill: { color: ACCENT } });
  s.addShape("rect", { x: 0, y: 4.1, w: 10, h: 1.525, fill: { color: SHADOW_DARK } });

  s.addText("One Toolkit.", {
    x: 0.8, y: 0.5, w: 8.8, h: 1.0, fontSize: 52, bold: true, color: WHITE, fontFace: "Calibri", margin: 0,
  });
  s.addText("Endless Capabilities.", {
    x: 0.8, y: 1.45, w: 8.8, h: 1.0, fontSize: 52, bold: true, color: ACCENT, fontFace: "Calibri", margin: 0,
  });
  s.addText("17 production-grade skills spanning creative design, software development,\nenterprise communication, document processing, and specialized workflows.", {
    x: 0.8, y: 2.6, w: 8.8, h: 1.1, fontSize: 16, color: ACCENT, fontFace: "Calibri", margin: 0,
  });

  [["17", "Skills"], ["5", "Categories"], ["Open", "Source"]].forEach(([n, l], i) => {
    const x = 0.8 + i * 3.0;
    s.addText(n,  { x, y: 4.22, w: 2.8, h: 0.52, fontSize: 28, bold: true, color: WHITE,  fontFace: "Calibri", margin: 0 });
    s.addText(l,  { x, y: 4.72, w: 2.8, h: 0.3,  fontSize: 13, color: ACCENT, fontFace: "Calibri", margin: 0 });
  });
}

// ── Write file ────────────────────────────────────────────────────────────────
pres.writeFile({ fileName: "claude-agent-skills.pptx" })
  .then(() => console.log("Done: claude-agent-skills.pptx"))
  .catch(e => { console.error(e); process.exit(1); });
