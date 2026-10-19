// Builds Beam_Executive_Briefing.docx: the written companion to the executive PowerPoint.
const fs = require("fs");
const path = require("path");
const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, Footer, Header, PageNumber, LevelFormat, PageBreak } = require("docx");
const { p, runs, b, h1, h2, h3, bullet, num, note, link, img, table, gap, BLUE, INK2 } = require("./docx_helpers");

const ROOT = path.resolve(__dirname, "..");
const children = [];
const add = (...xs) => xs.flat().forEach((x) => children.push(x));

// ---------- Cover ----------
add(
  new Paragraph({ spacing: { before: 1800, after: 160 }, children: [new TextRun({ text: "EXECUTIVE BRIEFING", size: 22, bold: true, color: "F2A516" })] }),
  new Paragraph({ spacing: { after: 200 }, children: [new TextRun({ text: "Building a Beam-Class LLM", size: 56, bold: true, color: BLUE })] }),
  new Paragraph({ spacing: { after: 500 }, children: [new TextRun({ text: "Should we develop a new open model that competes with Anthropic, OpenAI, Google, xAI and Meta? Findings, options, investment and the decision requested.", size: 28 })] }),
  table(["Item", "Detail"], [
    ["Prepared for", "Yossi Avni and the executive / investment committee"],
    ["Date", "19 October 2026"],
    ["Read with", "Beam_Executive_Summary.pptx (16 slides) — this briefing follows the same storyline"],
    ["Evidence base", "Beam_Research_Baseline.docx v2.0 (37 pages, 42 sources)"],
    ["Reading time", "≈ 15 minutes"],
    ["Decision requested", "Approve Stage 1 (≈ US$11M; first model in 3–6 months) and the choice of target market"],
  ], [2600, 6760]),
  gap(),
  note("Estimates of cost, GPU-hours and headcount are the author’s, derived from Reflection AI’s disclosed training scale. Benchmark figures are vendor-reported unless described as independent."),
);

// ---------- 1. One-page summary ----------
add(h1("1. One-page summary"),
  h3("Situation"),
  p("Reflection AI released Beam 1.0 on 16 October 2026: a 501-billion-parameter open-weight model (23.4B active per token) under the permissive Apache 2.0 licence, built for coding and agentic work and launched with 40+ partners including AWS, Google Cloud, Microsoft Azure and NVIDIA. It shows that a well-funded start-up can train a credible large model in under two years, and it makes a strong, free starting point available to anyone."),
  h3("Complication"),
  p("Beam is efficient but not frontier-class. On the only independent single-scale comparison (Artificial Analysis Intelligence Index) it scores 30 — versus 44–46 for the best Chinese open models and Grok 4.7, ~53 for OpenAI’s GPT-6 Astra and Google’s Gemini 4 Argon, and 58 for Anthropic’s Claude Opus 5.5. Meanwhile enterprises are concentrating spend on closed models: Anthropic (40%) and OpenAI (27%) lead, and open-weight models’ share fell from 19% to 11% in a year."),
  h3("Resolution"),
  num(runs(b("Do not race the leaderboard. "), "Matching Anthropic/OpenAI/Google requires US$1B+ per model generation.")),
  num(runs(b("Build where closed models cannot follow: "), "sovereign and on-premises AI, regulated industries, domain specialists, under-served languages, and cost-sensitive agent workloads.")),
  num(runs(b("Start from Beam’s open weights (Stage 1, ≈ US$11M). "), "Prove the data, reinforcement-learning and evaluation capability and win first customers within 6 months.")),
  num(runs(b("Gate further investment. "), "Only at month 9, with proven demand and ≥ US$150M secured, commit to continued pre-training (≈ US$56M) or a from-scratch model (≈ US$175M).")),
  h3("Key numbers"),
  table(["30 vs 58", "US$11M → US$175M", "11%", "3–6 months"], [
    ["Beam vs. leader, independent index", "Adapt Beam → build from scratch (2-yr, indicative)", "Open-weight share of enterprise API spend", "Time to a first adapted model"],
  ], [2340, 2340, 2340, 2340]),
);

// ---------- 2. Situation ----------
add(h1("2. What Reflection AI has built"),
  table(["Attribute", "Beam 1.0"], [
    ["Size", "501.3B total parameters; 23.4B active per token (≈4.7%)"],
    ["Design", "Mixture-of-Experts: 256 experts, 8 used per token; 52 layers; 512-token sliding-window attention with full-context attention in only 9 layers; 1M-token context"],
    ["Training", "23.8T tokens on 6,144 NVIDIA GB300 GPUs (< 4 weeks), then reinforcement learning on 10,500 GB300 GPUs for ~4 weeks across ~1M coding/agent tasks"],
    ["Release", "Apache 2.0 weights in BF16, FP8 and NVFP4; speculative-decoding drafter; Mirror CLI agent harness; day-0 support in vLLM, SGLang and Transformers"],
    ["Company", "Founded 2024 by ex-Google DeepMind researchers; US$2B raised Oct 2025 led by NVIDIA; reported valuation talks at ~US$25B in 2026"],
  ], [2000, 7360]),
  ...img("fig1_architecture.png", 6.3, "Beam 1.0 architecture (from the research baseline, Figure 1)."),
  p(runs(b("Why it matters to us: "), "every major design choice in Beam is a published technique, and its weights can be used commercially. The barrier to entry for a capable open model has fallen from hundreds of millions of dollars to roughly ten million for an adapted model.")),
);

// ---------- 3. Competitive position ----------
add(h1("3. How Beam compares with the leading vendors"),
  ...img("fig8_aa_index.png", 6.3, "Independent ranking on one scale (Artificial Analysis, October 2026)."),
  table(["Vendor", "Flagship", "Open weights?", "Independent index", "Strategic posture"], [
    ["Anthropic", "Claude Opus 5.5 / Sonnet 5.5", "No", "58 / 56", "Capability leader; largest enterprise share"],
    ["OpenAI", "GPT-6 Astra", "No (gpt-oss open line)", "≈ 53", "Consumer reach; premium pricing"],
    ["Google", "Gemini 4 Argon / 3.8 Flash", "No (Gemma open line)", "52.6", "Multimodal; cloud distribution; cheap fast tier"],
    ["xAI", "Grok 4.7", "No", "46", "Low price ($2/$6 per 1M tokens) and speed"],
    ["Meta", "Muse Spark (closed); Llama 4 (open)", "Partly", "n/a on same index", "Distribution; frontier models now closed"],
    ["Chinese labs", "MiMo, GLM, DeepSeek, Kimi", "Yes", "44–46", "Best open scores at the lowest prices"],
    ["Reflection", "Beam 1.0", "Yes (Apache 2.0)", "30", "US provenance; efficiency; ecosystem"],
  ], [1400, 2300, 1600, 1500, 2560]),
  gap(),
  p(runs(b("Reading the gap: "), "Reflection’s own launch table placed Beam level with GLM-5.2. Independent testing did not confirm this. The lesson for us is to measure every claim — ours and competitors’ — on our own independent harness.")),
  p(runs(b("Beam’s real strengths: "), "it reaches answers with ~40% fewer output tokens than the median model, runs on a single 8-GPU server in 4-bit precision with ≤ 1.7 points quality loss, and is free to modify and resell.")),
);

// ---------- 4. Opportunity ----------
add(h1("4. Where a new model can win"),
  ...img("fig10_market_share.png", 5.8, "Enterprise LLM API spend share (Menlo Ventures, late 2025)."),
  table(["Segment", "Why closed vendors are weak here", "What we would need"], [
    ["Sovereign & government", "Weights cannot leave vendor clouds; Chinese models are often barred", "US/EU/allied provenance, on-prem packaging, security accreditation"],
    ["Regulated industries", "Data-residency and audit requirements", "Private deployment, audit logs, model documentation"],
    ["Vertical specialists", "General models are not trained on proprietary workflows", "Domain data and RL environments from design partners"],
    ["Under-served languages", "Low commercial priority for global vendors", "Targeted data (e.g., Hebrew, Arabic, regional European languages)"],
    ["High-volume agents", "Per-token pricing makes scale expensive", "Low serving cost; Beam’s token efficiency"],
  ], [2200, 3500, 3660]),
  gap(),
  p(runs(b("Recommendation: "), "choose one primary and one secondary segment before Stage 1 starts. The choice drives data acquisition, evaluation design and the first customers.")),
);

// ---------- 5. Options and investment ----------
add(h1("5. Options and investment"),
  ...img("fig11_options_cost.png", 6.0, "Indicative two-year cost by route (author estimates)."),
  table(["", "A. Adapt Beam", "B. Continue pre-training", "C. Build from scratch"], [
    ["What", "Fine-tune and reinforcement-learn Beam for our segment", "Add 3–6T domain tokens to Beam, then full post-training", "Own architecture, tokenizer and 24T+ token training run"],
    ["2-year cost", "≈ US$11M", "≈ US$56M", "≈ US$175M"],
    ["Time to first model", "3–6 months", "9–12 months", "15–18 months"],
    ["Team", "15–25", "40–60", "80–120"],
    ["Expected capability", "Beam + domain gains", "Beam + 5–10 points in domain", "Beam parity; path to next tier with more capital"],
    ["IP", "Apache 2.0 derivative", "Substantially modified derivative", "Fully owned"],
    ["Main risk", "Limited differentiation", "Data cost; forgetting", "Execution and capital"],
  ], [1600, 2500, 2600, 2660]),
  h2("5.1 Stage 1 budget (option A, indicative)"),
  table(["Line item", "US$M", "Notes"], [
    ["Compute", "3", "≈ 0.5–1M GPU-hours for fine-tuning, RL and evaluation"],
    ["People", "6", "15–25 staff (RL, data, infra, evaluation, product) over the programme"],
    ["Data and RL environments", "1", "Licensing, annotation, design-partner environments"],
    ["Infrastructure, evaluation, safety", "1", "Serving stack, red-teaming, third-party evaluation"],
    ["Total", "≈ 11", "Excludes go-to-market spend"],
  ], [3000, 1400, 4960]),
  h2("5.2 Higher ambition levels"),
  table(["Level", "Independent target", "Indicative cost", "Who realistically plays"], [
    ["1. Beam parity", "≈ 30", "US$11–175M", "Start-ups, sovereign programmes, enterprises"],
    ["2. Open-frontier parity", "≈ 45", "US$300–500M", "Well-funded labs (Chinese labs, Reflection’s next model)"],
    ["3. Closed-frontier challenger", "≥ 53", "US$1B+ per generation", "Anthropic, OpenAI, Google, xAI, Meta"],
  ], [2400, 1800, 2000, 3160]),
);

// ---------- 6. Roadmap ----------
add(h1("6. Roadmap and decision gates"),
  ...img("fig7_roadmap.png", 6.3, "Programme roadmap; Stage 1 covers phases 0–2 and the adapted-model release."),
  table(["Gate", "Month", "Go criteria", "Decision"], [
    ["G0 Baseline", "1", "Beam and peers reproduced on our harness within ±2 points", "Confirm Stage 1 plan"],
    ["G1 Adapt", "6", "Adapted model +5 points on domain evaluations; first paying pilot", "Continue / pivot segment"],
    ["G2 Scale", "9", "Demand proven; pilot model meets scaling targets; ≥ US$150M secured", "Fund option B or C, or stay at A"],
    ["G3 Pre-train", "11", "Stable full training run", "Continue / stop"],
    ["G4 Release", "17", "Level 1 KPIs met independently; safety report; ≥ 10 partners", "Launch"],
  ], [1500, 900, 4400, 2560]),
);

// ---------- 7. Risks ----------
add(h1("7. Principal risks"),
  table(["Risk", "Likelihood", "Impact", "Mitigation"], [
    ["Frontier keeps moving away", "High", "High", "Compete in segments, not on the general leaderboard"],
    ["Chinese open models at near-zero prices", "High", "Medium", "Sell provenance, sovereignty, support and domain quality"],
    ["Our own claims fail to replicate", "Medium", "High", "Independent harness and third-party evaluation from day one"],
    ["GPU cost / availability", "Medium", "High", "Multi-provider reservations; 4-bit (NVFP4) training and serving"],
    ["RL infrastructure underestimated", "High", "High", "Build environments and sandboxes from month 1"],
    ["Talent and legal exposure", "Medium", "High", "Equity and academic partnerships; data-provenance tracking; legal gates"],
    ["Reflection ownership change", "Medium", "Medium", "Mirror Beam 1.0 weights now — the Apache 2.0 grant is irrevocable"],
  ], [2900, 1200, 1100, 4160]),
);

// ---------- 8. Decision & 90 days ----------
add(h1("8. Decision requested and first 90 days"),
  h2("8.1 Decisions requested"),
  num(runs(b("Approve Stage 1 (option A): "), "≈ US$11M programme budget; first adapted model in 3–6 months.")),
  num(runs(b("Select the target segment: "), "recommended primary — sovereign / regulated on-premises AI; secondary — one vertical or language market chosen with design partners.")),
  num(runs(b("Authorise partnerships: "), "one GPU cloud for compute reservations and two to three design-partner customers.")),
  num(runs(b("Schedule Gate 2 review "), "for month 9 to decide on options B or C.")),
  h2("8.2 First 90-day plan"),
  table(["Weeks", "Workstream", "Deliverable"], [
    ["1–2", "Programme set-up", "Mirror Beam 1.0 weights; GPU reservation; core hires (RL lead, data lead, infra lead)"],
    ["1–4", "Evaluation", "Independent harness; Beam, GLM, Kimi, DeepSeek re-run on all KPI benchmarks (Gate 0)"],
    ["2–8", "Market", "Segment selection; 2–3 design-partner agreements; domain evaluation sets"],
    ["3–10", "Data", "Domain corpus acquisition and filtering; synthetic data pipeline"],
    ["4–12", "RL environments", "First 10,000 verifiable domain tasks running in sandboxes on Mirror CLI"],
    ["8–12", "First experiments", "Fine-tuned and RL-trained Beam checkpoint; first domain results vs. base Beam"],
    ["12", "Review", "90-day report to the committee: results, spend, pipeline"],
  ], [1200, 2300, 5860]),
);

// ---------- 9. Presentation guide ----------
const slides = [
  ["1", "Building a Beam-Class LLM", "Frame the decision: whether and how to build a new open model against the leading vendors."],
  ["2", "The answer up front", "Three numbers: 30 vs 58 (gap), US$11–175M (cost range), 11% (open share). Verdict: buildable, but win on control and cost, not rank."],
  ["3", "Beam 1.0: what Reflection AI shipped", "501B / 23.4B active, 256 experts, 1M context, Apache 2.0. Released 16 Oct with 40+ partners."],
  ["4", "Architecture built for cheap tokens", "Sparse experts and sparse long attention make it cheap to serve; nothing proprietary — good news for us."],
  ["5", "Training: post-training is the big bill", "RL used ~2–3× the GPU-hours of pre-training. Environments, not just GPUs, are the long pole."],
  ["6", "On its own numbers, Beam trails open peers", "Vendor table: level with GLM-5.2, behind Kimi K3 and Qwen 3.8 Max."],
  ["7", "Independent tests: a 28-point gap", "The key slide. Independent index: Beam 30, open leaders 44–46, closed frontier 53–58."],
  ["8", "The vendor landscape", "Closed vendors compete on capability; Chinese labs on price; we need a third reason to exist."],
  ["9", "Enterprise spend favours closed models", "Anthropic 40%, OpenAI 27%; open share down to 11%. Growth is real but goes to the top models."],
  ["10", "Where a new open model can win", "Six segments; ask the room which one or two we can credibly own."],
  ["11", "Three levels of ambition", "Level 1 now; Level 2 is a funded option; Level 3 is a hyperscaler bet."],
  ["12", "Three routes to a Beam-class model", "A ≈ US$11M, B ≈ US$56M, C ≈ US$175M. Recommend A first."],
  ["13", "Staged roadmap with decision gates", "Gate 2 at month 9 is the real investment decision."],
  ["14", "Key risks and how we manage them", "Biggest risks are strategic (moving frontier, price war) and RL infrastructure."],
  ["15", "Recommendation and next steps", "Ask for approval of Stage 1 and the segment choice."],
  ["16", "Sources and notes", "Point to the 37-page baseline for evidence."],
];
add(h1("9. Presenter guide to the executive deck"),
  p("Suggested timing: 25 minutes presentation + 20 minutes discussion. Slides 2, 7, 12 and 15 carry the argument; the others can be shortened if time is tight. Speaker notes with the same points are embedded in the PowerPoint file."),
  table(["Slide", "Title", "Key message / talking point"], slides, [800, 3000, 5560]),
);

// ---------- Appendix ----------
add(h1("Appendix. Key sources"),
  ...[
    ["Reflection — Beam is now available", "https://reflection.ai/blog/beam-is-available"],
    ["Hugging Face — Beam 1.0 model card", "https://huggingface.co/reflection/Beam-1.0-501B-A23B"],
    ["Artificial Analysis — Reflection Beam", "https://artificialanalysis.ai/articles/reflection-beam"],
    ["Artificial Analysis — Intelligence Index", "https://artificialanalysis.ai/evaluations/artificial-analysis-intelligence-index"],
    ["OrcaRouter — Beam 1.0 scored 30 independently", "https://www.orcarouter.ai/blog/reflection-beam-1-0-open-weights-available"],
    ["AWS — Introducing Reflection Beam", "https://aws.amazon.com/blogs/machine-learning/introducing-reflection-beam-an-open-weight-model-for-agentic-coding-on-aws/"],
    ["llm-stats — Claude Opus 5.5 launch", "https://llm-stats.com/blog/research/claude-opus-5-5-launch"],
    ["The Decoder — Grok 4.7", "https://the-decoder.com/xai-launches-grok-4-7-at-bargain-prices-but-benchmarks-reveal-a-wide-gap-to-claude-and-gpt-6/"],
    ["Hostinger — LLM statistics 2026 (Menlo Ventures data)", "https://www.hostinger.com/tutorials/llm-statistics/"],
    ["BenchLM — LLM API pricing (Oct 2026)", "https://benchlm.ai/llm-pricing"],
  ].map(([t, u]) => bullet([link(t, u)])),
  note("Full source list (42 references) and all technical detail: Beam_Research_Baseline.docx."),
);

const doc = new Document({
  creator: "Beam research programme", title: "Building a Beam-Class LLM — Executive Briefing",
  styles: {
    default: { document: { run: { font: "Calibri", size: 22 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 34, bold: true, color: BLUE, font: "Calibri" }, paragraph: { spacing: { before: 240, after: 160 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, color: "1F3B63", font: "Calibri" }, paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 23, bold: true, color: "1F3B63", font: "Calibri" }, paragraph: { spacing: { before: 160, after: 80 }, outlineLevel: 2 } },
    ],
  },
  numbering: { config: [
    { reference: "bul", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 270 } } } }] },
    { reference: "num", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 360 } } } }] },
    { reference: "src", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "[%1]", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 600, hanging: 600 } } } }] },
  ] },
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } },
    headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "Building a Beam-Class LLM — Executive Briefing", size: 16, color: INK2 })] })] }) },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: ["Page ", PageNumber.CURRENT], size: 16, color: INK2 })] })] }) },
    children,
  }],
});
Packer.toBuffer(doc).then((buf) => {
  const out = path.join(ROOT, "Beam_Executive_Briefing.docx");
  fs.writeFileSync(out, buf);
  console.log("wrote", out);
});
