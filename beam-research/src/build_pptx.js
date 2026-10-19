// Builds Beam_Executive_Summary.pptx: an executive deck on the opportunity to build a Beam-class LLM.
// Run: NODE_PATH=<dir with pptxgenjs, react-icons, react, react-dom, sharp> node src/build_pptx.js <pptx-skill-dir>
const path = require("path");
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");

const SKILL = process.argv[2];
const { applyTheme } = require(path.join(SKILL, "scripts/apply_theme.js"));
const ROOT = path.resolve(__dirname, "..");
const OUT = path.join(ROOT, "Beam_Executive_Summary.pptx");

const THEME = {
  name: "Beam Executive",
  headFontFace: "Arial",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "0B1530", lt1: "FFFFFF", dk2: "13254F", lt2: "EEF2F8",
    accent1: "F2A516", accent2: "2E6FD8", accent3: "5BB4E5", accent4: "8A93A6",
    accent5: "1BAF7A", accent6: "E34948", hlink: "2E6FD8", folHlink: "6250D6",
  },
};
const HX = THEME.colors; // hex values for hex-only options (charts, shadows, gridlines)

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9"; // 10 x 5.625 in
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "Building a Beam-Class LLM — Executive Summary";
pres.author = "Beam research programme";
const C = pres.SchemeColor;

const FOOT = "Building a Beam-Class LLM · Executive summary · October 2026";
// ---------- layouts ----------
pres.defineSlideMaster({
  title: "TITLE_DARK",
  background: { color: C.text2 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 1.55, w: 8.8, h: 1.2, fontSize: 40, bold: true, color: C.background1, valign: "bottom", align: "left", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: 0.6, y: 2.9, w: 8.2, h: 1.1, fontSize: 18, color: C.background2, valign: "top", margin: 0 }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: C.background1 },
  margin: [0.5, 0.5, 0.6, 0.5],
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.5, y: 0.3, w: 9.0, h: 0.65, fontSize: 28, bold: true, color: C.text2, valign: "middle", margin: 0 }, text: "" } },
    { text: { text: FOOT, options: { x: 0.5, y: 5.25, w: 7.5, h: 0.25, fontSize: 9, color: C.accent4, margin: 0 } } },
  ],
  slideNumber: { x: 9.1, y: 5.25, w: 0.4, h: 0.25, fontSize: 9, color: C.accent4, align: "right" },
});
pres.defineSlideMaster({
  title: "CONTENT_DARK",
  background: { color: C.text2 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.5, y: 0.3, w: 9.0, h: 0.65, fontSize: 28, bold: true, color: C.background1, valign: "middle", margin: 0 }, text: "" } },
    { text: { text: FOOT, options: { x: 0.5, y: 5.25, w: 7.5, h: 0.25, fontSize: 9, color: C.accent3, margin: 0 } } },
  ],
  slideNumber: { x: 9.1, y: 5.25, w: 0.4, h: 0.25, fontSize: 9, color: C.accent3, align: "right" },
});

// ---------- helpers ----------
async function icon(Comp, hex) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Comp, { color: "#" + hex, size: 256 }));
  const png = await sharp(Buffer.from(svg)).resize(256, 256).png().toBuffer();
  return "image/png;base64," + png.toString("base64");
}
function iconBadge(slide, data, x, y, d, name) {
  slide.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: C.accent1 }, line: { color: C.accent1 }, objectName: name + " badge" });
  const pad = d * 0.24;
  slide.addImage({ data, x: x + pad, y: y + pad, w: d - 2 * pad, h: d - 2 * pad, objectName: name + " icon" });
}
const card = (slide, x, y, w, h, name, fill = C.background2) =>
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08, fill: { color: fill }, line: { color: fill }, objectName: name });
const txt = (slide, text, opts) => slide.addText(text, { isTextBox: true, margin: 0, valign: "top", ...opts });
const chartBase = {
  catAxisLabelColor: "52514E", valAxisLabelColor: "52514E", catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt",
  dataLabelFontFace: "+mn-lt", legendFontFace: "+mn-lt", titleFontFace: "+mn-lt",
  catAxisLabelFontSize: 11, valAxisLabelFontSize: 10, dataLabelFontSize: 10, legendFontSize: 11,
  valGridLine: { color: "E3E6EC", size: 0.75 }, catGridLine: { style: "none" }, dataLabelColor: "0B1530",
};

(async () => {
  const ic = {
    sovereign: await icon(fa.FaLandmark, HX.dk2), regulated: await icon(fa.FaShieldAlt, HX.dk2),
    vertical: await icon(fa.FaFlask, HX.dk2), cost: await icon(fa.FaCoins, HX.dk2), lang: await icon(fa.FaLanguage, HX.dk2),
    data: await icon(fa.FaDatabase, HX.dk2), rocket: await icon(fa.FaRocket, HX.dk2), cogs: await icon(fa.FaCogs, HX.dk2),
    warn: await icon(fa.FaExclamationTriangle, HX.dk2), users: await icon(fa.FaUsers, HX.dk2), server: await icon(fa.FaServer, HX.dk2),
    scale: await icon(fa.FaBalanceScale, HX.dk2),
  };

  // ===== 1. Title =====
  pres.addSection({ title: "Opening" });
  let s = pres.addSlide({ masterName: "TITLE_DARK", sectionTitle: "Opening" });
  s.addText("Building a Beam-Class LLM", { placeholder: "title" });
  s.addText("The opportunity to develop a new open model that competes with Anthropic, OpenAI, Google, xAI and Meta", { placeholder: "body" });
  txt(s, "Executive summary · 19 October 2026 · Companion to the Research Baseline v2.0", { x: 0.6, y: 4.7, w: 8.5, h: 0.3, fontSize: 12, color: C.accent3 });
  s.addShape(pres.shapes.OVAL, { x: 8.1, y: 0.5, w: 1.3, h: 1.3, fill: { color: C.accent1 }, line: { color: C.accent1 }, objectName: "beam motif" });
  s.addNotes("Purpose: decide whether and how to build a new open LLM comparable to Reflection AI's Beam, and how it could compete with the leading vendors. Detailed evidence is in the 37-page research baseline document.");

  // ===== 2. Answer up front =====
  s = pres.addSlide({ masterName: "CONTENT_DARK", sectionTitle: "Opening" });
  s.addText("The answer up front", { placeholder: "title" });
  const stats = [
    ["30 vs 58", "Beam vs. Claude Opus 5.5 on the independent Artificial Analysis Intelligence Index"],
    ["$11–175M", "Indicative 2-year cost to field a Beam-class model: adapt Beam vs. build from scratch"],
    ["11%", "Open-weight share of enterprise LLM API spend, down from 19% a year earlier"],
  ];
  stats.forEach(([big, small], i) => {
    const x = 0.5 + i * 3.08;
    txt(s, big, { x, y: 1.25, w: 2.85, h: 0.9, fontSize: 40, bold: true, color: C.accent1, fontFace: THEME.headFontFace, valign: "bottom" });
    txt(s, small, { x, y: 2.25, w: 2.75, h: 1.0, fontSize: 14, color: C.background1 });
  });
  card(s, 0.5, 3.6, 9.0, 1.35, "verdict card", C.text1);
  txt(s, [
    { text: "Verdict: ", options: { bold: true, color: C.accent1 } },
    { text: "a Beam-class model is buildable and affordable — but it will not beat Anthropic, OpenAI, Google or xAI on general capability. Win where closed models cannot follow: sovereignty, on-prem control, domain specialisation and cost. Start by adapting Beam’s Apache 2.0 weights; gate any from-scratch build on traction.", options: { color: C.background1 } },
  ], { x: 0.75, y: 3.75, w: 8.5, h: 1.1, fontSize: 15, valign: "middle" });
  s.addNotes("Three numbers frame the decision. First, Beam is 28 points behind the leader on the only single-scale independent index. Second, the cost range depends mainly on whether we start from Beam's open weights. Third, enterprises are concentrating spend on closed models, so an open model must win on dimensions other than rank.");

  // ===== 3. What is Beam =====
  pres.addSection({ title: "Beam today" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Beam today" });
  s.addText("Beam 1.0: what Reflection AI shipped", { placeholder: "title" });
  const specs = [["501B", "total parameters"], ["23.4B", "active per token"], ["256 / 8", "experts / active"],
    ["1M", "token context"], ["23.8T", "pre-training tokens"], ["Apache 2.0", "open weights, 16 Oct"]];
  specs.forEach(([v, l], i) => {
    const x = 0.5 + (i % 3) * 1.95, y = 1.15 + Math.floor(i / 3) * 1.55;
    card(s, x, y, 1.8, 1.35, "spec card " + (i + 1));
    txt(s, v, { x: x + 0.12, y: y + 0.18, w: 1.6, h: 0.6, fontSize: v.length > 6 ? 20 : 24, bold: true, color: C.accent2, fontFace: THEME.headFontFace, valign: "middle" });
    txt(s, l, { x: x + 0.12, y: y + 0.82, w: 1.6, h: 0.4, fontSize: 12, color: C.text1 });
  });
  txt(s, "Timeline", { x: 6.5, y: 1.15, w: 3.0, h: 0.35, fontSize: 16, bold: true, color: C.text2 });
  txt(s, [
    { text: "2024 · founded by ex-DeepMind team (Laskin, Antonoglou)", options: { bullet: true, breakLine: true } },
    { text: "Oct 2025 · $2B raise led by NVIDIA", options: { bullet: true, breakLine: true } },
    { text: "5 Oct 2026 · Beam announced", options: { bullet: true, breakLine: true } },
    { text: "16 Oct 2026 · weights, Mirror CLI harness and DFlash2 drafter released; 40+ partners", options: { bullet: true, breakLine: true } },
    { text: "Independent score: 30 (19th of 120)", options: { bullet: true } },
  ], { x: 6.5, y: 1.55, w: 3.0, h: 3.3, fontSize: 13, color: C.text1, paraSpaceAfter: 6 });
  s.addNotes("Beam is a sparse mixture-of-experts model: 501 billion parameters stored, about 23 billion used per token. It was released under Apache 2.0 with BF16, FP8 and NVFP4 checkpoints, a speculative-decoding drafter, and Mirror CLI, the agent harness it was trained against.");

  // ===== 4. Architecture =====
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Beam today" });
  s.addText("Architecture built for cheap tokens", { placeholder: "title" });
  s.addImage({ path: path.join(ROOT, "figures/fig1_architecture.png"), x: 0.5, y: 1.1, w: 5.95, h: 3.89, objectName: "architecture diagram" });
  const arch = [
    ["Sparse experts", "Top-8 of 256 experts: ~4.7% of weights active per token"],
    ["Sparse long attention", "512-token window in 43 layers; full context in only 9"],
    ["Ready to serve", "NVFP4 loses ≤1.7 pts; fits one 8-GPU node; drafter speeds agents"],
  ];
  arch.forEach(([h, b], i) => {
    const y = 1.15 + i * 1.3;
    txt(s, h, { x: 6.75, y, w: 2.75, h: 0.35, fontSize: 15, bold: true, color: C.accent2 });
    txt(s, b, { x: 6.75, y: y + 0.38, w: 2.75, h: 0.8, fontSize: 13, color: C.text1 });
  });
  s.addNotes("The design choices are all published techniques: DeepSeek-style fine-grained experts with auxiliary-loss-free balancing, a sliding-window/global attention mix with no positional encoding on global layers, gated attention, SandwichNorm and FP32 residuals. Nothing here is proprietary, which is good news for a replication.");

  // ===== 5. Training =====
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Beam today" });
  s.addText("Training: post-training is the big bill", { placeholder: "title" });
  const steps = [["Curate", "~95% of web text discarded"], ["Pre-train", "23.8T tokens · 6,144 GB300 · <4 wks"],
    ["Mid-train", "Context to 1M tokens"], ["SFT", "Tool use & agent traces"], ["RL", "10,500 GB300 · 4 wks · 100M+ rollouts"]];
  steps.forEach(([h, b], i) => {
    const x = 0.5 + i * 1.83;
    const hot = i === 4;
    card(s, x, 1.25, 1.65, 1.75, "step " + (i + 1), hot ? C.text2 : C.background2);
    txt(s, String(i + 1), { x: x + 0.15, y: 1.37, w: 0.5, h: 0.4, fontSize: 18, bold: true, color: C.accent1 });
    txt(s, h, { x: x + 0.15, y: 1.8, w: 1.4, h: 0.35, fontSize: 15, bold: true, color: hot ? C.background1 : C.text2 });
    txt(s, b, { x: x + 0.15, y: 2.18, w: 1.4, h: 0.75, fontSize: 12, color: hot ? C.background1 : C.text1 });
  });
  s.addChart(pres.charts.BAR, [{ name: "GB300 GPU-hours (millions, est.)", labels: ["Pre-training", "Reinforcement learning"], values: [2.5, 7.1] }], {
    x: 0.5, y: 3.2, w: 5.2, h: 1.9, barDir: "bar", chartColors: [HX.accent2], showValue: true, dataLabelPosition: "outEnd",
    dataLabelFormatCode: '0.0"M"', valAxisHidden: true, showLegend: false, showTitle: true, title: "Estimated GPU-hours by phase",
    titleFontSize: 12, titleColor: "13254F", ...chartBase, valGridLine: { style: "none" }, valAxisMaxVal: 9, objectName: "gpu hours chart",
  });
  txt(s, "RL used ~2–3× the GPU-hours of pre-training. Budget for an environment fleet (~1M tasks, ~1.3B sandboxes), not just a training cluster.",
    { x: 6.0, y: 3.45, w: 3.5, h: 1.5, fontSize: 14, color: C.text1 });
  s.addNotes("Pre-training compute is estimated at about 3.3e24 FLOPs. The RL phase ran on roughly 1.7 times as many GPUs for a similar duration. Building verifiable environments and sandbox orchestration is the longest-lead engineering task.");

  // ===== 6. Vendor-reported =====
  pres.addSection({ title: "Competition" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Competition" });
  s.addText("On its own numbers, Beam trails open peers", { placeholder: "title" });
  s.addChart(pres.charts.BAR, [
    { name: "Beam", labels: ["Terminal-Bench 2.1", "GPQA Diamond", "HLE (no tools)"], values: [80.1, 90.5, 36.2] },
    { name: "GLM 5.2", labels: ["Terminal-Bench 2.1", "GPQA Diamond", "HLE (no tools)"], values: [81.0, 91.2, 40.5] },
    { name: "Qwen 3.8 Max", labels: ["Terminal-Bench 2.1", "GPQA Diamond", "HLE (no tools)"], values: [86.6, 92.6, 43.6] },
    { name: "Kimi K3", labels: ["Terminal-Bench 2.1", "GPQA Diamond", "HLE (no tools)"], values: [88.3, 93.5, 46.9] },
  ], {
    x: 0.5, y: 1.1, w: 6.0, h: 3.95, barDir: "col", barGapWidthPct: 60, chartColors: [HX.accent1, "8A93A6", "B7BECC", "13254F"],
    showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0.0", dataLabelFontSize: 9, showLegend: true, legendPos: "b", valAxisMaxVal: 100, valAxisMinVal: 0,
    ...chartBase, dataLabelFontSize: 9, objectName: "vendor benchmarks chart",
  });
  txt(s, [
    { text: "Reflection’s launch table (vendor-reported)", options: { bold: true, breakLine: true } },
    { text: "Beam is roughly level with GLM-5.2 on agentic coding and science QA, but behind Kimi K3 and Qwen 3.8 Max on every shared benchmark.", options: { breakLine: true } },
    { text: " ", options: { breakLine: true } },
    { text: "Its pitch is efficiency: similar quality with 23B active parameters vs. 40–95B for peers.", options: {} },
  ], { x: 6.8, y: 1.2, w: 2.7, h: 3.8, fontSize: 14, color: C.text1 });
  s.addNotes("All figures here come from Reflection's own launch table. SWE-bench Verified is 81.6 per the model card, and SWE-bench Pro v1 is 66.1, where Beam beats GLM-5.2 (62.1).");

  // ===== 7. Independent ranking =====
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Competition" });
  s.addText("Independent tests: a 28-point gap", { placeholder: "title" });
  const aaNames = ["Claude Opus 5.5", "Claude Sonnet 5.5", "GPT-6 Astra", "Claude Fable 5.1", "Gemini 4 Argon", "Grok 4.7",
    "MiMo-V2.6-Pro", "GLM-5.3 Max", "DeepSeek V4.1 Flash", "Kimi K3", "Beam 1.0", "Nemotron 3.5 Super"];
  const aaVals = [58, 56, 53, 53, 52.6, 46, 46, 45, 45, 44, 30, 26];
  const rev = (a) => [...a].reverse();
  s.addChart(pres.charts.BAR, [
    { name: "Closed", labels: rev(aaNames), values: rev(aaVals.map((v, i) => (i < 6 ? v : 0))) },
    { name: "Open weights", labels: rev(aaNames), values: rev(aaVals.map((v, i) => (i >= 6 && i !== 10 ? v : 0))) },
    { name: "Beam", labels: rev(aaNames), values: rev(aaVals.map((v, i) => (i === 10 ? v : 0))) },
  ], {
    x: 0.5, y: 1.05, w: 5.9, h: 4.05, barDir: "bar", barGrouping: "stacked", barGapWidthPct: 35,
    chartColors: ["13254F", "8A93A6", HX.accent1], showValue: true, dataLabelPosition: "inEnd", dataLabelColor: "FFFFFF",
    dataLabelFormatCode: "#;;", showLegend: true, legendPos: "b", valAxisMaxVal: 60, valAxisHidden: true,
    ...chartBase, dataLabelColor: "FFFFFF", catAxisLabelFontSize: 10, valGridLine: { style: "none" }, objectName: "independent index chart",
  });
  txt(s, [
    { text: "Artificial Analysis Intelligence Index, Oct 2026", options: { bold: true, breakLine: true } },
    { text: "One scale, one harness, 120 models.", options: { breakLine: true } },
    { text: " ", options: { breakLine: true } },
    { text: "Closed frontier: 53–58", options: { bullet: true, breakLine: true } },
    { text: "Grok and top Chinese open models: 44–46", options: { bullet: true, breakLine: true } },
    { text: "Beam: 30, but ~40% fewer output tokens than the median", options: { bullet: true } },
  ], { x: 6.7, y: 1.2, w: 2.8, h: 3.8, fontSize: 14, color: C.text1, paraSpaceAfter: 4 });
  s.addNotes("This is the most important slide. Vendor tables put Beam near GLM-5.2, but the only single-scale independent index puts it 14 to 16 points behind the best open models and 23 to 28 behind Anthropic, OpenAI and Google. Meta's closed Muse Spark 1.3 scored 61 on an earlier index version, which is not comparable and is excluded here.");

  // ===== 8. Landscape table =====
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Competition" });
  s.addText("The vendor landscape", { placeholder: "title" });
  const hdr = (t) => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.text2 } } });
  const rows = [
    [hdr("Vendor"), hdr("Flagship"), hdr("Weights"), hdr("AA Index"), hdr("Output $/1M"), hdr("Edge")],
    ["Anthropic", "Claude Opus 5.5", "Closed", "58", "$20", "Top capability; 40% enterprise share"],
    ["OpenAI", "GPT-6 Astra", "Closed", "53", "$50", "Reach; reasoning"],
    ["Google", "Gemini 4 Argon", "Closed", "52.6", "—", "Multimodal; cloud; cheap Flash tier"],
    ["xAI", "Grok 4.7", "Closed", "46", "$6", "Price and speed"],
    ["Meta", "Muse Spark / Llama 4", "Closed / open", "—", "—", "Distribution; Llama brand"],
    ["China open labs", "MiMo · GLM · Kimi", "Open", "44–46", "$0.28+", "Best open scores; lowest prices"],
    [{ text: "Reflection", options: { bold: true } }, { text: "Beam 1.0", options: { bold: true } }, "Open", { text: "30", options: { bold: true } }, "Partner", "US provenance; efficiency"],
  ];
  s.addTable(rows, { x: 0.5, y: 1.1, w: 9.0, colW: [1.4, 1.9, 1.05, 0.9, 0.95, 2.8], fontSize: 12, color: C.text1,
    border: { type: "solid", pt: 0.5, color: "D5DAE3" }, rowH: 0.4, valign: "middle", fontFace: THEME.bodyFontFace, objectName: "landscape table" });
  txt(s, "Weights “Open” for Reflection = Apache 2.0. Output prices per 1M tokens (aggregators, Oct 2026). Groq (inference hardware) hosts open models and is a deployment partner, not a model competitor.",
    { x: 0.5, y: 4.75, w: 9.0, h: 0.4, fontSize: 10, color: C.accent4, italic: true });
  s.addNotes("Closed vendors compete on capability; Chinese open labs compete on price and open scores. A new US- or EU-built open model sits between them and needs a clear reason to be chosen.");

  // ===== 9. Market =====
  pres.addSection({ title: "Opportunity" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Opportunity" });
  s.addText("Enterprise spend favours closed models", { placeholder: "title" });
  s.addChart(pres.charts.BAR, [{ name: "Share of enterprise LLM API spend (%)", labels: ["Anthropic", "OpenAI", "Google & other closed", "Open-weight models"], values: [40, 27, 22, 11] }], {
    x: 0.5, y: 1.1, w: 5.4, h: 3.0, barDir: "bar", chartColors: [HX.accent2], showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: '0"%"',
    showLegend: false, valAxisHidden: true, valAxisMaxVal: 50, ...chartBase, valGridLine: { style: "none" }, objectName: "market share chart",
  });
  txt(s, "Source: Menlo Ventures, late 2025. “Google & other closed” is the remainder.", { x: 0.5, y: 4.15, w: 5.4, h: 0.3, fontSize: 10, color: C.accent4, italic: true });
  const mk = [["$37B", "enterprise gen-AI spend in 2025 (≈3× y/y)"], ["19% → 11%", "open-weight share in one year"]];
  mk.forEach(([v, l], i) => {
    const y = 1.15 + i * 1.6;
    card(s, 6.3, y, 3.2, 1.4, "market stat " + (i + 1));
    txt(s, v, { x: 6.45, y: y + 0.15, w: 2.9, h: 0.6, fontSize: 26, bold: true, color: C.accent2, fontFace: THEME.headFontFace });
    txt(s, l, { x: 6.45, y: y + 0.8, w: 2.9, h: 0.5, fontSize: 13, color: C.text1 });
  });
  txt(s, "The market is growing fast, but buyers pay for the most capable closed models. An open model has to offer something closed APIs cannot.",
    { x: 0.5, y: 4.5, w: 9.0, h: 0.6, fontSize: 14, color: C.text2, bold: true });
  s.addNotes("Market-size estimates vary by definition; use them directionally. The consistent signal across sources is that open-weight share of enterprise spend has fallen.");

  // ===== 10. Where we can win =====
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Opportunity" });
  s.addText("Where a new open model can win", { placeholder: "title" });
  const wins = [
    [ic.sovereign, "Sovereign & government", "Allies that cannot procure Chinese weights and need on-prem control"],
    [ic.regulated, "Regulated industries", "Finance, health, defence, legal: residency and audit beat a few index points"],
    [ic.vertical, "Vertical specialists", "Post-train on proprietary environments to beat general models on narrow tasks"],
    [ic.cost, "High-volume agents", "Coding and automation agents where cost per token dominates"],
    [ic.lang, "Under-served languages", "Hebrew, Arabic, regional European languages via targeted data"],
    [ic.scale, "Customisation", "Full fine-tuning, RL on customer tasks, predictable pricing"],
  ];
  wins.forEach(([im, h, b], i) => {
    const x = 0.5 + (i % 3) * 3.05, y = 1.15 + Math.floor(i / 3) * 1.95;
    card(s, x, y, 2.85, 1.8, "win card " + (i + 1));
    iconBadge(s, im, x + 0.15, y + 0.15, 0.55, h);
    txt(s, h, { x: x + 0.82, y: y + 0.2, w: 1.95, h: 0.5, fontSize: 14, bold: true, color: C.text2, valign: "middle" });
    txt(s, b, { x: x + 0.15, y: y + 0.82, w: 2.6, h: 0.9, fontSize: 12, color: C.text1 });
  });
  s.addNotes("These are the segments where open weights win despite a lower aggregate score. Pick one or two as the commercial wedge; this choice should drive data, RL environments and evaluations.");

  // ===== 11. Ambition levels =====
  pres.addSection({ title: "Plan" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Plan" });
  s.addText("Three levels of ambition", { placeholder: "title" });
  const lv = [
    ["Level 1", "Beam parity", "AA ≈ 30", "~500B total / 23B active; Beam KPIs", "$11M–$175M", C.background2, C.text2],
    ["Level 2", "Open-frontier parity", "AA ≈ 45", "1–2T total / 40–60B active; 30T+ tokens", "$300–500M", C.background2, C.text2],
    ["Level 3", "Frontier challenger", "AA ≥ 53", "Multi-trillion params; 10^25–10^26 FLOPs", "$1B+ per generation", C.text2, C.background1],
  ];
  lv.forEach(([tag, name, aa, scale, cost, fill, ink], i) => {
    const x = 0.5 + i * 3.05;
    card(s, x, 1.15, 2.85, 3.55, "level card " + (i + 1), fill);
    txt(s, tag, { x: x + 0.2, y: 1.3, w: 2.5, h: 0.3, fontSize: 12, bold: true, color: C.accent1 });
    txt(s, name, { x: x + 0.2, y: 1.6, w: 2.5, h: 0.45, fontSize: 18, bold: true, color: ink });
    txt(s, aa, { x: x + 0.2, y: 2.15, w: 2.5, h: 0.55, fontSize: 26, bold: true, color: i === 2 ? C.accent1 : C.accent2, fontFace: THEME.headFontFace });
    txt(s, scale, { x: x + 0.2, y: 2.85, w: 2.5, h: 0.75, fontSize: 13, color: ink });
    txt(s, [{ text: "2-yr cost: ", options: { bold: true } }, { text: cost }], { x: x + 0.2, y: 3.9, w: 2.5, h: 0.6, fontSize: 13, color: ink });
  });
  txt(s, "Recommendation: commit to Level 1 now; treat Level 2 as a funded option after proof of demand. Level 3 is a hyperscaler-scale bet.",
    { x: 0.5, y: 4.78, w: 9.0, h: 0.4, fontSize: 13, color: C.text2, bold: true });
  s.addNotes("Competing head-on with Anthropic, OpenAI, Google or xAI requires Level 3 resources. The realistic path to relevance is Level 1 with a clear niche, then Level 2 if the business case is proven.");

  // ===== 12. Options and cost =====
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Plan" });
  s.addText("Three routes to a Beam-class model", { placeholder: "title" });
  const optL = ["A. Adapt Beam", "B. Continue pre-training", "C. Build from scratch"];
  s.addChart(pres.charts.BAR, [
    { name: "Compute", labels: optL, values: [3, 18, 55] },
    { name: "People (2 yrs)", labels: optL, values: [6, 25, 80] },
    { name: "Data & environments", labels: optL, values: [1, 8, 25] },
    { name: "Infra, eval, safety", labels: optL, values: [1, 5, 15] },
  ], {
    x: 0.5, y: 1.05, w: 5.6, h: 4.05, barDir: "col", barGrouping: "stacked", barGapWidthPct: 70,
    chartColors: ["13254F", HX.accent2, "5BB4E5", HX.accent1], showValue: false, showLegend: true, legendPos: "b", valAxisTitle: "US$ millions (indicative)", showValAxisTitle: true,
    valAxisTitleFontSize: 10, valAxisTitleColor: "52514E", ...chartBase, dataLabelColor: "FFFFFF", dataLabelFontSize: 9, objectName: "options cost chart",
  });
  const opt = [["A · ≈$11M · 3–6 months", "SFT + RL on Beam weights for a niche; 15–25 people"],
    ["B · ≈$56M · 9–12 months", "Add 3–6T domain tokens + full RL; 40–60 people"],
    ["C · ≈$175M · 15–18 months", "Own model, ~11–14M GPU-hours; 80–120 people"]];
  opt.forEach(([h, b], i) => {
    const y = 1.15 + i * 1.3;
    txt(s, h, { x: 6.4, y, w: 3.1, h: 0.35, fontSize: 14, bold: true, color: C.text2 });
    txt(s, b, { x: 6.4, y: y + 0.38, w: 3.1, h: 0.75, fontSize: 13, color: C.text1 });
  });
  s.addNotes("Author estimates at US$3–5 per GPU-hour and US$300–450K per fully loaded head-year. Option A gives the fastest proof of the data, RL and evaluation stack, and its weights are a permissive Apache 2.0 derivative.");

  // ===== 13. Roadmap =====
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Plan" });
  s.addText("Staged roadmap with decision gates", { placeholder: "title" });
  s.addShape(pres.shapes.LINE, { x: 0.8, y: 2.35, w: 8.4, h: 0, line: { color: C.accent4, width: 2 }, objectName: "timeline axis" });
  const gates = [["Month 1", "G0 Baseline", "Reproduce Beam & peers on our own harness"],
    ["Month 6", "G1 Adapt", "Adapted Beam +5 pts on domain evals; first paying pilot"],
    ["Month 9", "G2 Scale?", "Pilot model hits scaling targets; 20T tokens; ≥$150M funding"],
    ["Month 11", "G3 Pre-train", "Stable full training run under way"],
    ["Month 17", "G4 Release", "Level 1 KPIs met independently; safety report; partners"]];
  gates.forEach(([m, g, d], i) => {
    const x = 0.5 + i * 1.83;
    s.addShape(pres.shapes.OVAL, { x: x + 0.62, y: 2.17, w: 0.36, h: 0.36, fill: { color: i === 2 ? C.accent1 : C.accent2 }, line: { color: C.background1, width: 2 }, objectName: "gate " + (i + 1) });
    txt(s, m, { x, y: 1.55, w: 1.6, h: 0.3, fontSize: 12, color: C.accent4, align: "center" });
    txt(s, g, { x, y: 1.85, w: 1.6, h: 0.3, fontSize: 14, bold: true, color: C.text2, align: "center" });
    card(s, x, 2.8, 1.6, 1.3, "gate detail " + (i + 1));
    txt(s, d, { x: x + 0.12, y: 2.92, w: 1.36, h: 1.4, fontSize: 12, color: C.text1 });
  });
  txt(s, "G2 is the key investment decision: only proceed to a from-scratch model if Stage 1 shows demand and funding is secured.",
    { x: 0.5, y: 4.35, w: 9.0, h: 0.5, fontSize: 13, bold: true, color: C.text2 });
  s.addNotes("Each gate has explicit go criteria. Stage 1 (adapt Beam) de-risks the data engine, RL environments and evaluation harness before committing to large compute.");

  // ===== 14. Risks =====
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Plan" });
  s.addText("Key risks and how we manage them", { placeholder: "title" });
  const risks = [
    [ic.warn, "Frontier keeps moving", "Don’t race the leaderboard; own a niche"],
    [ic.cost, "Chinese open models at near-zero price", "Sell provenance, sovereignty, support"],
    [ic.scale, "Claims that don’t replicate", "Independent harness from day one"],
    [ic.server, "GPU cost and availability", "Multi-provider reservations; NVFP4"],
    [ic.cogs, "RL infrastructure underestimated", "Build environments from month 1"],
    [ic.users, "Talent and legal exposure", "Equity, partnerships; provenance tracking"],
  ];
  risks.forEach(([im, h, b], i) => {
    const x = 0.5 + (i % 2) * 4.6, y = 1.12 + Math.floor(i / 2) * 1.32;
    card(s, x, y, 4.4, 1.17, "risk card " + (i + 1));
    iconBadge(s, im, x + 0.18, y + 0.3, 0.55, h);
    txt(s, h, { x: x + 0.9, y: y + 0.17, w: 3.4, h: 0.38, fontSize: 14, bold: true, color: C.text2 });
    txt(s, b, { x: x + 0.9, y: y + 0.57, w: 3.4, h: 0.45, fontSize: 13, color: C.text1 });
  });
  s.addNotes("A full risk register with likelihood and impact is in Section 20 of the research baseline. Also mirror Beam 1.0 weights now: the Apache 2.0 grant is irrevocable for the released version, whatever happens to Reflection's ownership.");

  // ===== 15. Recommendation =====
  pres.addSection({ title: "Close" });
  s = pres.addSlide({ masterName: "CONTENT_DARK", sectionTitle: "Close" });
  s.addText("Recommendation and next steps", { placeholder: "title" });
  const nx = [["1", "Choose the wedge", "Pick one or two segments: sovereign, regulated, vertical or language"],
    ["2", "Adapt Beam (Stage 1)", "≈$11M, 3–6 months: SFT + RL on Beam 1.0 for that wedge"],
    ["3", "Build the hard parts early", "Data engine, RL environment fleet, independent eval harness"],
    ["4", "Decide at Gate 2", "Fund continued pre-training or a from-scratch model only on proven demand"]];
  nx.forEach(([n, h, b], i) => {
    const y = 1.15 + i * 0.95;
    s.addShape(pres.shapes.OVAL, { x: 0.5, y, w: 0.6, h: 0.6, fill: { color: C.accent1 }, line: { color: C.accent1 }, objectName: "step badge " + n });
    txt(s, n, { x: 0.5, y, w: 0.6, h: 0.6, fontSize: 20, bold: true, color: C.text2, align: "center", valign: "middle" });
    txt(s, h, { x: 1.35, y: y - 0.02, w: 8.1, h: 0.35, fontSize: 17, bold: true, color: C.background1 });
    txt(s, b, { x: 1.35, y: y + 0.33, w: 8.1, h: 0.35, fontSize: 14, color: C.background2 });
  });
  s.addNotes("Decision requested: approve Stage 1 budget (about US$11M over 3–6 months) and the choice of wedge market. Revisit the larger investment at Gate 2.");

  // ===== 16. Sources =====
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Close" });
  s.addText("Sources and notes", { placeholder: "title" });
  txt(s, [
    { text: "Reflection AI: Introducing Beam (5 Oct 2026); Beam is now available (16 Oct 2026)", options: { bullet: true, breakLine: true } },
    { text: "Hugging Face model cards: reflection/Beam-1.0-501B-A23B (BF16, FP8, NVFP4); vLLM and NVIDIA serving recipes", options: { bullet: true, breakLine: true } },
    { text: "Artificial Analysis Intelligence Index v4.3.x and Beam article (Oct 2026)", options: { bullet: true, breakLine: true } },
    { text: "Vendor launch materials and system cards: Anthropic (Claude Opus 5.5, Sonnet 5.5, Fable 5.1), OpenAI (GPT-6), Google (Gemini), xAI (Grok 4.7), Meta (Muse, Llama 4)", options: { bullet: true, breakLine: true } },
    { text: "Menlo Ventures enterprise LLM market data; pricing aggregators (BenchLM, Oct 2026)", options: { bullet: true, breakLine: true } },
    { text: "Costs, GPU-hours and team sizes are author estimates; see Sections 12 and Appendix A of the research baseline", options: { bullet: true } },
  ], { x: 0.5, y: 1.15, w: 9.0, h: 3.0, fontSize: 13, color: C.text1, paraSpaceAfter: 6 });
  card(s, 0.5, 3.75, 9.0, 0.8, "caveat card");
  txt(s, "Caveat: vendors report different benchmark versions and harnesses. Only the Artificial Analysis index puts all models on one scale; treat cross-vendor gaps under ~2 points as noise. Full source list (42 references) in the research baseline document.",
    { x: 0.7, y: 3.82, w: 8.6, h: 0.66, fontSize: 12, color: C.text1, valign: "middle" });
  s.addNotes("Full citations with links are in Section 24 of Beam_Research_Baseline.docx.");

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
