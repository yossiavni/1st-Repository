// Builds Beam_Research_Baseline.docx from the research notes and generated figures.
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, Table, TableRow, TableCell,
  WidthType, ShadingType, ImageRun, PageBreak, TableOfContents, Footer, Header, PageNumber,
  LevelFormat, ExternalHyperlink, BorderStyle,
} = require("docx");

const ROOT = path.resolve(__dirname, "..");
const FIG = (f) => path.join(ROOT, "figures", f);
const BLUE = "2A78D6", INK2 = "52514E";
const W = 9360; // content width in DXA (US Letter, 1" margins)

const p = (text, opts = {}) => new Paragraph({ spacing: { after: 120 }, ...opts,
  children: Array.isArray(text) ? text : [new TextRun(text)] });
const runs = (...parts) => parts.map((x) => typeof x === "string" ? new TextRun(x) : new TextRun(x));
const b = (t) => ({ text: t, bold: true });
const h1 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, pageBreakBefore: true, children: [new TextRun(t)] });
const h2 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(t)] });
const h3 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun(t)] });
const bullet = (t, level = 0) => new Paragraph({ numbering: { reference: "bul", level }, spacing: { after: 60 },
  children: Array.isArray(t) ? t : [new TextRun(t)] });
const num = (t) => new Paragraph({ numbering: { reference: "num", level: 0 }, spacing: { after: 60 },
  children: Array.isArray(t) ? t : [new TextRun(t)] });
const note = (t) => new Paragraph({ spacing: { after: 160 }, children: [new TextRun({ text: t, italics: true, size: 18, color: INK2 })] });
const link = (text, url) => new ExternalHyperlink({ link: url, children: [new TextRun({ text, style: "Hyperlink" })] });

function img(file, wIn, caption) {
  const buf = fs.readFileSync(FIG(file));
  const wpx = buf.readUInt32BE(16), hpx = buf.readUInt32BE(20);
  const w = wIn * 96, h = Math.round(w * hpx / wpx);
  return [
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 120, after: 60 },
      children: [new ImageRun({ type: "png", data: buf, transformation: { width: w, height: h } })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 },
      children: [new TextRun({ text: caption, italics: true, size: 18, color: INK2 })] }),
  ];
}

const border = { style: BorderStyle.SINGLE, size: 4, color: "D0CFC8" };
const borders = { top: border, bottom: border, left: border, right: border };
function table(headers, rows, widths) {
  const total = widths.reduce((a, c) => a + c, 0);
  const cell = (t, i, hdr) => new TableCell({
    borders, width: { size: widths[i], type: WidthType.DXA },
    shading: hdr ? { fill: "DBE9F9", type: ShadingType.CLEAR, color: "auto" } : undefined,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    children: [new Paragraph({ children: [new TextRun({ text: String(t), bold: hdr, size: 18 })] })],
  });
  return new Table({
    width: { size: total, type: WidthType.DXA }, columnWidths: widths,
    rows: [new TableRow({ tableHeader: true, children: headers.map((h, i) => cell(h, i, true)) }),
      ...rows.map((r) => new TableRow({ children: r.map((c, i) => cell(c, i, false)) }))],
  });
}
const gap = () => new Paragraph({ spacing: { after: 120 }, children: [] });

const children = [];
const add = (...xs) => xs.flat().forEach((x) => children.push(x));

// ---------------- Title page ----------------
add(
  new Paragraph({ spacing: { before: 2400, after: 200 }, children: [new TextRun({ text: "Reflection AI Beam", size: 60, bold: true, color: BLUE })] }),
  new Paragraph({ spacing: { after: 200 }, children: [new TextRun({ text: "Research Baseline for Building a Beam-Class Open MoE Language Model", size: 34 })] }),
  new Paragraph({ spacing: { after: 600 }, children: [new TextRun({ text: "Published details · architecture · benchmarks vs. Anthropic, OpenAI, Google, Meta and open-weight peers · KPI targets · recommendations · download links", size: 22, color: INK2 })] }),
  p(runs(b("Prepared for: "), "Yossi Avni")),
  p(runs(b("Date: "), "14 October 2026")),
  p(runs(b("Information cut-off: "), "public sources available up to 13–14 October 2026 (nine days after Beam's announcement)")),
  p(runs(b("Status of Beam weights at cut-off: "), "announced, not yet published (Apache 2.0 promised for October 2026)")),
  gap(),
  note("Reading guide: every Beam number in this document is vendor-reported unless marked otherwise. Items marked “estimate” or “*” are analyst or author inferences, not Reflection disclosures. Re-check Section 11 links and the Beam technical report once released; several architecture fields are expected to be filled in by it."),
  new Paragraph({ children: [new PageBreak()] }),
  new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("Contents")] }),
  new TableOfContents("Contents", { hyperlink: true, headingStyleRange: "1-2" }),
  note("Right-click the table of contents in Word and choose “Update Field” to refresh page numbers."),
);

// ---------------- 1. Executive summary ----------------
add(h1("1. Executive summary"),
  p("Reflection AI announced Beam on 5 October 2026: its first open-weight model, a sparse Mixture-of-Experts (MoE) transformer with 501 billion total parameters and about 23 billion active per token (“501B-A23B”). It is text-only, aimed at coding, reasoning and agentic workloads, was pretrained from scratch on 23.8 trillion tokens, extended to a 1-million-token context, and then post-trained with an unusually large reinforcement-learning (RL) run on 10,500 NVIDIA GB300 GPUs for about four weeks."),
  p("Its headline claim is efficiency rather than raw leadership: Reflection says Beam matches Z.ai’s GLM-5.2 on advanced reasoning with 3–4× less inference compute. On Reflection’s own table Beam trails Kimi K3 and Qwen 3.8 Max on most rows, is roughly at parity with GLM-5.2, and sits clearly behind the closed frontier (Claude Fable 5.1, GPT-6-series, Gemini 3.x) on the hardest benchmarks such as Humanity’s Last Exam."),
  h3("Key takeaways for a team building a Beam-equivalent model"),
  bullet(runs(b("The target is achievable at mid-scale. "), "Beam is the smallest model in its competitive set (Figure 3). A ~500B-A20–25B MoE trained on 20–25T curated tokens plus heavy agentic RL is the right reference design.")),
  bullet(runs(b("RL compute is the differentiator. "), "Beam’s RL phase used ~1.7× the GPUs of pretraining for a similar duration, across ~1M environments and >100M rollouts. Budget for post-training compute ≥ pretraining compute.")),
  bullet(runs(b("Data quality over quantity. "), "Reported curation discarded ~95% of raw web tokens. A data engine (filtering, synthetic code/math, executable environments) is the longest-lead item in the plan.")),
  bullet(runs(b("Architecture is mainstream 2026 practice. "), "Interleaved sliding-window/global attention, fine-grained experts, auxiliary-loss-free balancing, SandwichNorm, gated attention and FP32 residuals are all published techniques; none is proprietary.")),
  bullet(runs(b("Budget estimate: "), "~10M GB300-hours in total (≈2–3M pretraining + ≈7M RL), on the order of US$30–60M of compute at market rates, plus data and staff (Section 9).")),
  note("Terminology: the request mentioned “PKI”; this document interprets it as KPI (Key Performance Indicators) — the benchmark and efficiency targets a new model must reach. Section 8 defines them."),
);

// ---------------- 2. Reflection AI ----------------
add(h1("2. Reflection AI: company background"),
  table(["Item", "Detail"], [
    ["Founded", "2024, by former Google DeepMind researchers"],
    ["Founders", "Misha Laskin (CEO); Ioannis Antonoglou (CTO, co-creator of AlphaGo)"],
    ["Mission", "An American open frontier lab: open-weight models competitive with leading Chinese open models (DeepSeek, Qwen, Kimi, GLM)"],
    ["Funding", "US$2B round Oct 2025 led by NVIDIA (~US$800M), at US$8B valuation; reported talks in Mar 2026 for US$2.5B at ~US$25B pre-money; one aggregator lists US$27.5B (May 2026, unconfirmed)"],
    ["First product", "Asimov — a code-research agent (waitlist in 2025–26)"],
    ["First model", "Beam, announced 5 Oct 2026; weights promised under Apache 2.0 in October 2026"],
    ["Recent news", "Reports (Oct 2026) of early talks in which NVIDIA could acquire, invest further in, or license technology from Reflection — a reason to mirror weights promptly once released"],
  ], [2200, 7160]),
  gap(),
  note("Do not confuse with “Reflection Llama-3.1 70B” (2024, by Matt Shumer / HyperWrite) — an unrelated fine-tune with a similar name."),
);

// ---------------- 3. Published specs ----------------
add(h1("3. Beam: everything published so far"),
  table(["Attribute", "Value", "Confidence / source"], [
    ["Model type", "Sparse Mixture-of-Experts decoder-only transformer", "Reflection launch post"],
    ["Total parameters", "501B", "Reflection"],
    ["Active parameters / token", "~23B (≈4.6% of total)", "Reflection"],
    ["Layers", "52", "Secondary coverage"],
    ["Attention", "Interleaved local (sliding-window) and global attention", "Reflection; 3:1 ratio is an analyst reading"],
    ["Attention gating", "Elementwise gating on attention output", "Reflection"],
    ["Normalisation", "SandwichNorm (pre- and post-sublayer norm)", "Reflection"],
    ["Residual stream", "Depth-based residual scaling; FP32 residual accumulation", "Reflection"],
    ["Experts", "Fine-grained routed experts, learned top-k router; counts not disclosed", "Reflection"],
    ["Load balancing", "Auxiliary-loss-free (expert-bias) balancing with cosine decay of bias updates; busiest expert ≈1.04× mean load", "Reflection"],
    ["Modality", "Text only", "DataCamp, TrendingTopics"],
    ["Pretraining data", "23.8T tokens; ~95% of raw web tokens discarded in curation", "Reflection; curation figure from MarkTechPost"],
    ["Pretraining compute", "6,144 NVIDIA GB300 GPUs, < 4 weeks", "Reflection / coverage"],
    ["Context window", "1M tokens after mid-training; 262,144 tokens in beta API", "MarkTechPost; Kingy AI"],
    ["RL phase", "10,500 GB300 GPUs, ~4 weeks, >100M rollouts, ~1M coding/agentic/STEM environments, ~1.3B sandboxes, 256K max rollout context", "Reflection; MarkTechPost"],
    ["Efficiency claim", "Matches GLM-5.2 on reasoning at 3–4× lower inference compute (excludes prefill and serving overhead)", "Reflection; RuntimeWire caveat"],
    ["License", "Apache 2.0 (promised)", "Reflection"],
    ["Availability (14 Oct 2026)", "Waitlisted beta API; weights, model card and technical report not yet published", "OrcaRouter, OpenModelWeights"],
    ["Pricing", "No public rate card; not yet on Artificial Analysis index", "eesel.ai, OrcaRouter"],
    ["Not disclosed", "Tokenizer & vocabulary, hidden size, head counts, expert count / top-k / shared experts, optimizer, MTP, precision recipe", "—"],
  ], [2300, 4360, 2700]),
);

// ---------------- 4. Architecture ----------------
add(h1("4. Architecture deep-dive"),
  ...img("fig1_architecture.png", 6.5, "Figure 1 — Beam architecture reconstructed from public disclosures. Starred items await the technical report."),
  h2("4.1 Sparse MoE backbone"),
  p("Every layer replaces the dense feed-forward block with a bank of many small (“fine-grained”) experts. A learned router scores all experts for each token and dispatches it to the top-k. Only ~23B of 501B parameters are touched per token, so per-token FLOPs resemble a ~23B dense model while capacity resembles a ~500B model. Fine-grained experts (DeepSeekMoE lineage) give the router more combinations and improve specialisation."),
  h2("4.2 Auxiliary-loss-free load balancing with cosine-decayed bias updates"),
  p("Instead of adding a balancing loss that fights the language-modelling objective, each expert carries a bias b_i that is added to its routing score only for top-k selection. After each step the bias is nudged up for under-loaded experts and down for over-loaded ones. Reflection’s contribution is to decay the bias-update step on a cosine schedule so routing stabilises late in training. Reported result: the busiest expert averages only 1.04× the mean load — near-perfect utilisation, which matters for expert-parallel throughput."),
  h2("4.3 Interleaved local / global attention"),
  p("Most layers use sliding-window (local) attention whose KV cache is bounded by the window; periodic global layers attend over the full sequence. This cuts KV-cache memory and long-context FLOPs roughly in proportion to the share of local layers, which is how a 1M-token context becomes affordable. The same pattern appears in Gemma 3, MiMo-V2.6 and other 2026 models."),
  h2("4.4 Stability components"),
  bullet(runs(b("SandwichNorm: "), "normalisation both before and after each sub-layer, bounding activation growth in deep, sparse networks.")),
  bullet(runs(b("Elementwise attention gating: "), "a learned sigmoid gate on the attention output; known to suppress attention sinks and loss spikes and to help long-context extrapolation.")),
  bullet(runs(b("Depth-based residual scaling: "), "scaling residual branch contributions by layer depth for stable signal propagation.")),
  bullet(runs(b("FP32 residual accumulation: "), "the residual stream is accumulated in FP32 even when matmuls run in low precision (BF16/FP8), avoiding drift over 52 layers.")),
);

// ---------------- 5. Training ----------------
add(h1("5. Training recipe"),
  ...img("fig2_pipeline.png", 6.5, "Figure 2 — Beam’s disclosed training pipeline."),
  table(["Phase", "Disclosed facts", "What it implies for a replication"], [
    ["Data curation", "Aggressive filtering (~95% of raw web discarded)", "Invest in classifiers, dedup, decontamination and synthetic code/math before training"],
    ["Pretraining", "23.8T tokens on 6,144 GB300 GPUs in < 4 weeks", "≈ 3.3×10^24 training FLOPs (6·N_active·D); needs ≈ 20–35% MoE MFU"],
    ["Mid-training", "Context extension to 1M tokens", "Staged RoPE/window extension with long-document and repository-level data"],
    ["SFT", "Not detailed", "Tool-use, multi-turn agent traces and reasoning formats seed the RL policy"],
    ["RL", "10,500 GB300, ~4 weeks, >100M rollouts, ~1M environments, ~1.3B sandboxes", "Asynchronous RL infrastructure, sandbox orchestration and verifiable rewards are core engineering"],
  ], [1700, 3600, 4060]),
  gap(),
  p(runs(b("Observation: "), "In GPU-hours Beam’s RL phase (≈7.1M GB300-hours) was roughly 2–3× its pretraining phase. This inverts the classic ratio and reflects the 2026 consensus that agentic capability is mostly bought in post-training.")),
);

// ---------------- 6. Benchmarks ----------------
add(h1("6. Performance vs. leading models"),
  h2("6.1 Model scale in context"),
  ...img("fig3_params.png", 6.5, "Figure 3 — Total and active parameters of Beam and its open-weight competitors."),
  table(["Model (developer)", "Total params", "Active params", "Open weights?"], [
    ["Beam (Reflection AI)", "501B", "~23B", "Promised, Apache 2.0"],
    ["GLM-5.2 (Z.ai / Zhipu)", "~753B", "~40B", "Yes"],
    ["Mistral Large 4 (Mistral AI)", "~1T", "~49B", "Per launch coverage"],
    ["DeepSeek V4 Pro (DeepSeek)", "~1.6T", "~49B", "Yes"],
    ["Qwen 3.8 Max (Alibaba)", "~2.4T", "~95B", "Base checkpoint; hosted Max adds features"],
    ["Kimi K3 (Moonshot AI)", "~2.8T", "~50B (disputed; some report 104B)", "Yes"],
    ["Claude Fable 5.1 / Opus 5.5 (Anthropic)", "Undisclosed", "Undisclosed", "No (API only)"],
    ["GPT-6 Astra / GPT-6.1 Sol (OpenAI)", "Undisclosed", "Undisclosed", "No (API only)"],
    ["Gemini 3.1 Pro / 3.8 Flash (Google)", "Undisclosed", "Undisclosed", "No (API only)"],
    ["Llama 4 Maverick (Meta)", "~400B", "~17B", "Yes (Llama license)"],
  ], [3300, 1700, 2360, 2000]),
  h2("6.2 Beam vs. open-weight peers (Reflection’s launch table)"),
  ...img("fig4_bench_open.png", 6.5, "Figure 4 — Reflection-reported scores; NR = not reported."),
  table(["Benchmark", "Beam", "GLM 5.2", "Kimi K3", "Qwen 3.8 Max"], [
    ["SWE-Bench Pro v1", "65.5", "62.1", "NR", "67.7"],
    ["SWE-Bench Pro v2-Hard", "77.2", "NR", "88.2", "NR"],
    ["SWE-Bench Verified", "80.9", "NR", "NR", "NR"],
    ["Terminal-Bench 2.1", "80.1", "81.0", "88.3", "86.6"],
    ["GPQA Diamond", "90.5", "91.2", "93.5", "92.6"],
    ["HLE (no tools)", "36.2", "40.5", "46.9", "43.6"],
    ["AIME 2026", "97.8", "99.2", "NR", "NR"],
  ], [3160, 1550, 1550, 1550, 1550]),
  gap(),
  p("Reading: Beam beats GLM-5.2 on SWE-Bench Pro v1 and is within ~1 point on Terminal-Bench and GPQA, but trails it by 4 points on HLE. Kimi K3 and Qwen 3.8 Max lead every row on which they report. Beam’s value proposition is therefore per-token efficiency (Figure 6), not absolute leadership."),
  h2("6.3 Beam vs. the closed frontier (Anthropic, OpenAI, Google) and Meta"),
  ...img("fig5_bench_frontier.png", 6.5, "Figure 5 — Best available public numbers; versions and dates differ across vendors."),
  table(["Model", "GPQA Diamond", "HLE (no tools)", "Terminal-Bench 2.1", "Notes"], [
    ["Beam (Reflection)", "90.5", "36.2", "80.1", "Vendor-reported, Oct 2026"],
    ["Claude Fable 5.1 (Anthropic)", "92.6", "60.9 (65.0 w/ tools)", "—", "SWE-bench Pro 81.2; Terminal-Bench 4.0 55.8 (Sep 2026 system card)"],
    ["Claude Opus 5.5 (Anthropic)", "—", "—", "—", "Comparable published scores not found; trails GPT-6.1 Sol by 2.2 pts on AutomationBench per one source"],
    ["GPT-6.1 Sol (OpenAI)", "—", "—", "—", "Terminal-Bench Science 57.0; OSWorld 2.0 71.4; DeepSWE 71.9–75.2"],
    ["GPT-5.4 Pro / GPT-5.5 (OpenAI)", "94.4 (5.4 Pro)", "~41.4 (5.5)", "—", "Earlier 2026 generation"],
    ["Gemini 3.8 Flash (Google)", "95.6", "54.9 (HLE-Verified)", "90.8 (89.4 alt.)", "Sep 2026; SWE-Bench Pro 61.6"],
    ["Gemini 3.1 Pro (Google)", "94.3", "44.4", "68.5 (v2.0)", "Feb 2026"],
    ["Llama 4 Maverick (Meta)", "69.8", "—", "—", "Apr 2025; latest open Llama; Meta’s new Muse line is closed"],
    ["Groq", "n/a", "n/a", "n/a", "Groq builds LPU inference hardware and hosts others’ open models; it does not train frontier LLMs"],
  ], [2300, 1200, 1500, 1500, 2860]),
  gap(),
  h3("How to read the frontier comparison"),
  bullet("Closed-model vendors increasingly publish newer benchmark versions (Terminal-Bench 4.0, SWE-bench Pro, OSWorld 2.0) while open-model vendors publish Terminal-Bench 2.1 and SWE-Bench Pro v1/v2, so direct like-for-like numbers are sparse. Gaps under ~2 points are within harness noise."),
  bullet("On the least-saturated benchmark (HLE, no tools) Beam’s 36.2 is ~25 points behind Claude Fable 5.1 and ~8 points behind Gemini 3.1 Pro — the clearest measure of the frontier gap."),
  bullet("GPQA Diamond is near saturation (90–96% for all frontier and top open models); it no longer separates models well."),
  bullet("Groq is relevant as a deployment target (very high tokens/s for open MoE models) rather than as a competitor. NVIDIA licensed Groq’s LPU technology in Dec 2025 (~US$20B)."),
  h2("6.4 Efficiency: reasoning per active parameter"),
  ...img("fig6_efficiency.png", 5.5, "Figure 6 — GPQA Diamond vs. active parameters per token (open-weight models)."),
  p("Beam reaches ~97% of Kimi K3’s GPQA score with under half of its active parameters, and ~98% of Qwen 3.8 Max’s with under a quarter. For serving cost, active parameters drive decode FLOPs while total parameters drive memory footprint; at 501B total, Beam fits on a single 8×GB300 node in FP8 (≈0.5 TB weights) — a practical advantage over 1.6–2.8T-parameter peers."),
);

// ---------------- 7. Gaps ----------------
add(h1("7. What is still unknown (open questions)"),
  table(["Unknown", "Why it matters for replication", "Expected source"], [
    ["Expert count, top-k, shared experts, expert width", "Determines active/total ratio, routing granularity and expert-parallel layout", "Technical report / config.json"],
    ["Hidden size, heads, KV heads, window size, global ratio", "Determines KV-cache size and long-context cost", "config.json"],
    ["Tokenizer & vocabulary", "Affects code/multilingual efficiency and compatibility", "Tokenizer files"],
    ["Optimizer (AdamW vs. Muon) and LR schedule", "Muon has shown ~2× token efficiency at scale", "Technical report"],
    ["Precision (BF16 / FP8 / NVFP4)", "Drives GB300 throughput and stability", "Technical report"],
    ["Data mixture proportions", "Biggest driver of coding/agentic quality", "Technical report (partial at best)"],
    ["RL algorithm & reward design", "GRPO/PPO variant, async off-policy correction, reward models vs. verifiers", "Technical report"],
    ["Independent benchmark results", "All scores are vendor-reported", "Artificial Analysis, LMArena, Vals, SWE-bench leaderboards"],
  ], [2900, 3800, 2660]),
);

// ---------------- 8. KPIs ----------------
add(h1("8. KPI targets for a Beam-equivalent model"),
  p("A new model is “Beam-equivalent” if it meets the following KPIs on an identical evaluation harness. Run Beam itself through the same harness once weights ship — that is the only fair baseline."),
  table(["KPI", "Beam reference", "Minimum target", "Stretch target"], [
    ["SWE-Bench Pro v1", "65.5", "≥ 65", "≥ 68 (Qwen 3.8 Max)"],
    ["SWE-Bench Verified", "80.9", "≥ 80", "≥ 83"],
    ["Terminal-Bench 2.1", "80.1", "≥ 80", "≥ 86"],
    ["GPQA Diamond", "90.5", "≥ 90", "≥ 92"],
    ["HLE (no tools)", "36.2", "≥ 36", "≥ 42"],
    ["AIME 2026", "97.8", "≥ 97", "≥ 99"],
    ["Active parameters / token", "~23B", "≤ 25B", "≤ 20B"],
    ["Total parameters", "501B", "≤ 550B", "≤ 400B"],
    ["Context window (effective)", "1M (262K served)", "≥ 256K with RULER ≥ 85%", "1M"],
    ["Decode throughput, 8×GB300 FP8", "Not published", "Measure Beam; match it", "+20%"],
    ["Inference compute vs. GLM-5.2 at equal quality", "3–4× lower", "≥ 3× lower", "≥ 4× lower"],
    ["Expert load imbalance (max / mean)", "1.04×", "≤ 1.10×", "≤ 1.05×"],
    ["License", "Apache 2.0", "Permissive open weights", "Apache 2.0 + open data card"],
  ], [3000, 2000, 2180, 2180]),
);

// ---------------- 9. Reference design & budget ----------------
add(h1("9. Proposed reference design and compute budget"),
  h2("9.1 Starting-point configuration (to be validated by ablation)"),
  table(["Component", "Proposed choice", "Rationale"], [
    ["Size", "~480–520B total, 20–24B active, ~52–60 layers", "Matches Beam; fits one 8-GPU node in FP8"],
    ["Experts", "256 routed fine-grained experts + 1 shared, top-8", "DeepSeek-V3/V4-style granularity; ablate 128/256/384"],
    ["Routing", "Aux-loss-free bias balancing with cosine-decayed update rate; sequence-level balance check", "Reproduces Beam’s 1.04× load"],
    ["Attention", "3:1 sliding-window (4–8K window) : global, GQA or MLA on global layers, gated output", "Bounded KV cache, 1M-context capable"],
    ["Norm / residual", "SandwichNorm (RMSNorm), depth-scaled residuals, FP32 residual accumulation, QK-norm", "Stability at low precision"],
    ["Position encoding", "RoPE on global layers with YaRN-style extension; NoPE/partial RoPE as ablation", "Long-context extrapolation"],
    ["Prediction head", "Multi-token prediction (1–2 extra heads)", "Denser training signal + speculative decoding"],
    ["Tokenizer", "~160–200K BPE vocabulary, code- and multilingual-aware", "Compression on code and non-English text"],
    ["Optimizer", "Muon for matrix params (per-expert), AdamW for embeddings/norms", "~2× token efficiency in published studies"],
    ["Precision", "FP8 (block-scaled) training, ablate NVFP4 on GB300", "Throughput on Blackwell-class GPUs"],
  ], [1800, 4100, 3460]),
  h2("9.2 Compute and cost estimate"),
  table(["Phase", "Estimate", "Assumptions"], [
    ["Pretraining", "≈ 3.3×10^24 FLOPs → ≈ 2–3M GB300-hours", "6 × 23B active × 23.8T tokens; 20–30% MFU on ~2.5 PFLOP/s BF16-dense (higher with FP8)"],
    ["Mid-training (long context)", "≈ 0.2–0.4M GB300-hours", "~1T tokens at growing sequence length"],
    ["RL", "≈ 7M GB300-hours", "10,500 GPUs × 28 days, as disclosed for Beam"],
    ["Ablations & scaling pilots", "≈ 1–2M GPU-hours", "1–7B and ~60B-A3B models"],
    ["Total", "≈ 10–12M GPU-hours", "≈ US$30–60M at US$3–5 per GPU-hour, excluding staff and data"],
  ], [2400, 3200, 3760]),
  note("These are author estimates derived from Beam’s disclosed GPU counts and durations, not Reflection figures."),
  h2("9.3 Program roadmap"),
  ...img("fig7_roadmap.png", 6.5, "Figure 7 — Proposed ~17-month program to reach Beam-equivalent KPIs."),
);

// ---------------- 10. Recommendations ----------------
add(h1("10. Recommendations from current research"),
  h2("10.1 Architecture"),
  num(runs(b("Adopt fine-grained MoE with aux-loss-free balancing from day one. "), "It is now standard (DeepSeek-V3/V4, Beam). Add Reflection’s cosine decay on bias updates and monitor max/mean expert load per layer.")),
  num(runs(b("Use hybrid attention for long context. "), "Interleave sliding-window and global layers (Beam, MiMo-V2.6); evaluate compressed/sparse global attention as in DeepSeek-V4 for 1M-token contexts, and dynamic sparse attention for decode.")),
  num(runs(b("Gate attention outputs and use SandwichNorm + QK-norm. "), "Cheap insurance against loss spikes during multi-trillion-token runs.")),
  num(runs(b("Add multi-token prediction. "), "Improves sample efficiency and gives a built-in speculative decoder (used by DeepSeek-V3/V4, MiMo-V2.6).")),
  h2("10.2 Optimisation and precision"),
  num(runs(b("Evaluate Muon. "), "Moonlight reported ~2× compute efficiency vs AdamW; DeepSeek-V4 trains with Muon and orthogonalises each expert separately; 2026 studies find it scales better with batch size.")),
  num(runs(b("Train in FP8, pilot NVFP4. "), "Blackwell/GB300 FP4 throughput is roughly double FP8; keep FP32 master weights and residual accumulation as Beam does.")),
  num(runs(b("Use WSD (warmup-stable-decay) learning-rate schedules "), "so that intermediate checkpoints can be annealed for evaluation and mid-training branches.")),
  h2("10.3 Data"),
  num(runs(b("Build the data engine first. "), "Model-based quality classifiers, aggressive dedup, benchmark decontamination, and synthetic code/math with execution-verified answers. Beam’s ~95% discard rate signals how strict this must be.")),
  num(runs(b("Collect repository-level and long-document data "), "for mid-training; include agent traces (terminal sessions, tool calls) for SFT.")),
  h2("10.4 Post-training and agentic RL"),
  num(runs(b("Plan RL compute ≥ pretraining compute. "), "Beam’s 10,500-GPU RL run and MiMo-V2.6’s RL-for-self-improvement framing both point the same way.")),
  num(runs(b("Invest in environments, not just prompts. "), "Target 10^5–10^6 verifiable environments (unit-test-graded repos, terminal tasks, math/STEM checkers, web/tool tasks) with a sandbox fleet that scales to billions of executions.")),
  num(runs(b("Use asynchronous, off-policy-corrected RL "), "(GRPO-family objectives with importance-ratio clipping) to keep inference and training GPUs saturated; support 128K–256K rollout context.")),
  h2("10.5 Evaluation"),
  num(runs(b("Evaluate on unsaturated, contamination-resistant suites: "), "SWE-bench Pro (OpenAI no longer uses SWE-bench Verified for frontier coding), Terminal-Bench 2.x/4.0, HLE, AIME of the current year, RULER/long-context, and agentic suites (OSWorld, τ-bench).")),
  num(runs(b("Report with confidence intervals and fixed harnesses; "), "re-run competitors yourself. Vendor-to-vendor gaps of 1–2 points are noise.")),
  h2("10.6 Strategy"),
  num(runs(b("Fine-tune before you pre-train. "), "Once Beam (Apache 2.0) ships, a Beam- or GLM-based fine-tune is the fastest way to validate your data, RL stack and evaluation harness before committing ~US$50M to pretraining from scratch.")),
  num(runs(b("Differentiate on one axis. "), "Beam’s angle is efficiency per active parameter; pick yours (e.g., domain specialisation, multimodality, on-device distillation, or a sub-10B-active model).")),
  num(runs(b("Mirror open checkpoints immediately on release "), "— Apache 2.0 grants are irrevocable, but hosting can change if ownership does.")),
);

// ---------------- 11. Downloads ----------------
add(h1("11. Download links and access"),
  p(runs(b("Beam status (14 Oct 2026): "), "weights are not yet downloadable. Reflection has promised Apache 2.0 weights, a model card and a technical report in October 2026; early access is via a waitlisted beta API.")),
  table(["Model / resource", "Where", "Status"], [
    ["Reflection Beam — announcement", "reflection.ai/blog/introducing-beam", "Live"],
    ["Reflection Beam — weights", "huggingface.co (search “reflection”)", "Pending (promised Oct 2026)"],
    ["DeepSeek V4 family", "huggingface.co/deepseek-ai", "Open weights"],
    ["Kimi K3 (Moonshot)", "huggingface.co/moonshotai", "Open weights"],
    ["GLM-5.x (Z.ai)", "huggingface.co/zai-org", "Open weights"],
    ["Qwen 3.x (Alibaba)", "huggingface.co/Qwen", "Open weights"],
    ["Mistral (Mistral AI)", "huggingface.co/mistralai", "Open weights (check license per model)"],
    ["Llama 4 (Meta)", "huggingface.co/meta-llama", "Gated, Llama license"],
    ["gpt-oss (OpenAI open-weight)", "huggingface.co/openai", "Open weights"],
    ["Claude (Anthropic)", "console.anthropic.com", "API only"],
    ["GPT-6 family (OpenAI)", "platform.openai.com", "API only"],
    ["Gemini (Google)", "aistudio.google.com", "API only"],
    ["GroqCloud (fast inference of open models)", "console.groq.com", "API (hosted open models)"],
  ], [3400, 3600, 2360]),
  gap(),
  p("Clickable links:"),
  bullet([link("Reflection — Introducing Beam", "https://reflection.ai/blog/introducing-beam")]),
  bullet([link("Hugging Face search: Reflection models", "https://huggingface.co/models?search=reflection")]),
  bullet([link("DeepSeek on Hugging Face", "https://huggingface.co/deepseek-ai")]),
  bullet([link("Moonshot AI (Kimi) on Hugging Face", "https://huggingface.co/moonshotai")]),
  bullet([link("Z.ai (GLM) on Hugging Face", "https://huggingface.co/zai-org")]),
  bullet([link("Qwen on Hugging Face", "https://huggingface.co/Qwen")]),
  bullet([link("Mistral AI on Hugging Face", "https://huggingface.co/mistralai")]),
  bullet([link("Meta Llama on Hugging Face", "https://huggingface.co/meta-llama")]),
  bullet([link("OpenAI on Hugging Face", "https://huggingface.co/openai")]),
  bullet([link("Anthropic Console", "https://console.anthropic.com")]),
  bullet([link("OpenAI Platform", "https://platform.openai.com")]),
  bullet([link("Google AI Studio", "https://aistudio.google.com")]),
  bullet([link("GroqCloud Console", "https://console.groq.com")]),
  note("Exact Beam repository IDs could not be verified before release; use the Hugging Face search link and confirm the publisher is Reflection’s verified organisation."),
);

// ---------------- 12. Sources ----------------
const sources = [
  ["Reflection — Introducing Beam (launch post)", "https://reflection.ai/blog/introducing-beam"],
  ["Reflection on X — Beam announcement", "https://x.com/reflection_ai/status/2107186849370247235"],
  ["TechCrunch — Reflection debuts Beam", "https://techcrunch.com/2026/10/05/reflection-debuts-beam-a-open-weight-ai-model-to-rival-chinese-models-at-lower-compute-cost/"],
  ["SiliconANGLE — Beam with 501B parameters", "https://siliconangle.com/2026/10/05/reflection-ai-debuts-open-source-beam-model-with-501b-parameters/"],
  ["Fortune — Reflection AI unveils Beam", "https://fortune.com/2026/10/05/reflection-ai-unveils-beam-a-new-us-based-open-source-model-to-compete-with-china/"],
  ["MarkTechPost — Beam 501B open-weight MoE", "https://www.marktechpost.com/2026/10/05/reflection-ai-introduces-beam-a-501b-open-weight-moe-model-with-23b-active-parameters-for-coding-and-agentic-workloads/"],
  ["DataCamp — Beam explained", "https://www.datacamp.com/blog/reflection-ai-beam"],
  ["MindStudio — Beam explained", "https://www.mindstudio.ai/blog/reflection-ai-beam-open-weight-model"],
  ["Latent Space AINews — Reflection Beam 501B-A23B", "https://www.latent.space/p/ainews-reflection-beam-501b-a23b"],
  ["Kingy AI — Specs, benchmarks & API limits", "https://kingy.ai/news/reflection-beam-501b-open-weight-ai-guide/"],
  ["Munder Difflin — Beam explained", "https://munderdiffl.in/blog/reflection-beam/"],
  ["RuntimeWire — Beam benchmark scores", "https://runtimewire.com/article/reflection-ai-beam-open-weight-model"],
  ["OrcaRouter — Weights not out", "https://www.orcarouter.ai/blog/reflection-beam"],
  ["eesel.ai — Specs, access and price", "https://www.eesel.ai/blog/reflection-ai-beam"],
  ["OpenModelWeights — Beam preview", "https://openmodelweights.com/news/reflection-beam-501b-open-weight-preview/"],
  ["Wikipedia — Reflection AI", "https://en.wikipedia.org/wiki/Reflection_AI"],
  ["Anthropic — Claude Fable 5.1 & Mythos 5.1 System Card", "https://www-cdn.anthropic.com/0339e6a7c5c7b87f5c07798616dc32c215d14235/Claude%20Fable%205.1%20&%20Claude%20Mythos%205.1%20System%20Card.pdf"],
  ["llm-stats — Claude Fable 5.1", "https://llm-stats.com/models/claude-fable-5-1"],
  ["Vellum — GPT-6.1 Sol benchmarks", "https://www.vellum.ai/blog/gpt-6-1-sol-benchmarks-explained"],
  ["BenchLM — GPT-6.1 Sol", "https://benchlm.ai/models/gpt-6-1-sol"],
  ["DataCamp — Gemini 3.8 Flash", "https://www.datacamp.com/blog/gemini-3-8-flash-cyber"],
  ["OpenRouter — GPQA Diamond leaderboard", "https://openrouter.ai/benchmarks/gpqa-diamond"],
  ["Digital Applied — Gemini 3.1 Pro benchmarks", "https://www.digitalapplied.com/blog/google-gemini-3-1-pro-benchmarks-pricing-guide"],
  ["ComputingForGeeks — Open-source LLM comparison 2026", "https://computingforgeeks.com/open-source-llm-comparison/"],
  ["DeepInfra — Kimi K3 vs DeepSeek V4 Pro vs GLM-5.2", "https://deepinfra.com/blog/kimi-k3-vs-deepseek-v4-pro-vs-glm-5-2"],
  ["Codersera — Llama 4 guide 2026", "https://codersera.com/blog/llama-4-complete-guide-2026/"],
  ["Voiceflow — Groq in 2026", "https://www.voiceflow.com/blog/groq"],
  ["DeepSeek-V4 technical report (arXiv)", "https://arxiv.org/html/2606.19348v1"],
  ["MiMo-V2.6: Scaling RL Towards Self-Improvement (arXiv)", "https://arxiv.org/html/2610.11959v1"],
  ["Muon is Scalable for LLM Training (arXiv)", "https://arxiv.org/pdf/2502.16982"],
  ["SOAP, Muon, and Beyond (arXiv)", "https://arxiv.org/pdf/2607.20548"],
  ["SWE-Bench Pro Verified (arXiv)", "https://arxiv.org/pdf/2609.08149"],
];
add(h1("12. Sources"),
  ...sources.map(([t, u]) => new Paragraph({ numbering: { reference: "src", level: 0 }, spacing: { after: 40 },
    children: [link(t, u)] })),
);

// ---------------- Appendix ----------------
add(h1("Appendix A. Glossary"),
  table(["Term", "Meaning"], [
    ["MoE", "Mixture-of-Experts: layers with many expert sub-networks, only a few active per token"],
    ["501B-A23B", "501B total parameters, 23B active per token"],
    ["Aux-loss-free balancing", "Load balancing via per-expert routing biases instead of an extra loss term"],
    ["SWA", "Sliding-window attention: each token attends only to a recent window"],
    ["GQA / MLA", "Grouped-query / multi-head latent attention: ways to shrink the KV cache"],
    ["SandwichNorm", "Normalisation before and after each sub-layer"],
    ["MTP", "Multi-token prediction: extra heads predicting tokens further ahead"],
    ["Muon", "Optimizer that orthogonalises momentum updates for matrix parameters"],
    ["MFU", "Model FLOPs Utilisation: achieved / peak hardware FLOPs"],
    ["GPQA Diamond", "Graduate-level science multiple-choice benchmark"],
    ["HLE", "Humanity’s Last Exam: expert-written questions at the frontier of knowledge"],
    ["SWE-Bench Pro / Verified", "Real GitHub issue-fixing benchmarks; Pro is harder and less contaminated"],
    ["Terminal-Bench", "Agentic tasks performed in a real terminal"],
    ["AIME", "American Invitational Mathematics Examination problems"],
  ], [2400, 6960]),
);

const doc = new Document({
  creator: "Beam research baseline", title: "Reflection AI Beam — Research Baseline",
  styles: {
    default: { document: { run: { font: "Calibri", size: 21 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 34, bold: true, color: BLUE, font: "Calibri" }, paragraph: { spacing: { before: 240, after: 160 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, color: "1F3B63", font: "Calibri" }, paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 22, bold: true, color: "1F3B63", font: "Calibri" }, paragraph: { spacing: { before: 160, after: 80 }, outlineLevel: 2 } },
    ],
  },
  numbering: { config: [
    { reference: "bul", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 540, hanging: 270 } } } }] },
    { reference: "num", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 540, hanging: 360 } } } }] },
    { reference: "src", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "[%1]", alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 600, hanging: 600 } } } }] },
  ] },
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } },
    headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT,
      children: [new TextRun({ text: "Reflection AI Beam — Research Baseline", size: 16, color: INK2 })] })] }) },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER,
      children: [new TextRun({ children: ["Page ", PageNumber.CURRENT], size: 16, color: INK2 })] })] }) },
    children,
  }],
});

Packer.toBuffer(doc).then((buf) => {
  const out = path.join(ROOT, "Beam_Research_Baseline.docx");
  fs.writeFileSync(out, buf);
  console.log("wrote", out);
});
