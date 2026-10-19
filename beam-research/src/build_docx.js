// Builds Beam_Research_Baseline.docx: the full research baseline and business case.
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, TableOfContents, Footer, Header,
  PageNumber, LevelFormat, PageBreak,
} = require("docx");
const { p, runs, b, h1, h2, h3, bullet, num, note, link, img, table, gap, BLUE, INK2 } = require("./docx_helpers");

const ROOT = path.resolve(__dirname, "..");
const children = [];
const add = (...xs) => xs.flat().forEach((x) => children.push(x));
const kv = (rows, w = [2600, 6760]) => table(["Item", "Detail"], rows, w);

// =============== Title & control ===============
add(
  new Paragraph({ spacing: { before: 2000, after: 200 }, children: [new TextRun({ text: "Building a Beam-Class LLM", size: 60, bold: true, color: BLUE })] }),
  new Paragraph({ spacing: { after: 200 }, children: [new TextRun({ text: "Reflection AI Beam 1.0: full technical analysis, competitive benchmarking against Anthropic, OpenAI, Google, xAI, Meta and open-weight leaders, and a business case for developing a competing model", size: 30 })] }),
  new Paragraph({ spacing: { after: 500 }, children: [new TextRun({ text: "Research baseline · Version 2.0", size: 24, color: INK2 })] }),
  kv([
    ["Prepared for", "Yossi Avni"],
    ["Version / date", "2.0, 19 October 2026 (supersedes v1.0 of 14 October 2026)"],
    ["Information cut-off", "Public sources available to 18–19 October 2026"],
    ["Companion deliverable", "Beam_Executive_Summary.pptx (executive deck)"],
    ["What changed in v2.0", "Beam weights released 16 Oct (Apache 2.0); model-card architecture confirmed (256 experts, top-8, 152K vocab, SWA-512); first independent score (Artificial Analysis Index = 30); xAI Grok, Meta Muse, Claude Opus/Sonnet 5.5, GPT-6, Gemini 4 added; market, pricing, business case, cost model, governance and risk sections added"],
    ["Confidence convention", "[V] vendor-reported · [I] independently measured · [E] author estimate · [S] single secondary source"],
  ]),
  gap(),
  note("All benchmark figures are as published by the cited sources and were not re-run by the author. Several frontier vendors report on different benchmark versions and harnesses; cross-vendor differences under ~2 points should be treated as noise."),
  new Paragraph({ children: [new PageBreak()] }),
  new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("Contents")] }),
  new TableOfContents("Contents", { hyperlink: true, headingStyleRange: "1-2" }),
  note("In Word: right-click the table of contents → Update Field → Update entire table."),
);

// =============== 1. Executive summary ===============
add(h1("1. Executive summary"),
  h2("1.1 The bottom line"),
  p("Reflection AI’s Beam 1.0 is a 501B-parameter sparse Mixture-of-Experts (MoE) model with 23.4B active parameters per token, released with open weights under Apache 2.0 on 16 October 2026. It is the most capable American open-weight model released from a start-up to date, built for coding and agentic work, and it is unusually efficient: it ships with a 512-token sliding-window attention design that keeps a full-length KV cache in only 9 of its 52 layers, and it uses ~40% fewer output tokens than comparable models on the Artificial Analysis test suite."),
  p("However, Beam is not a frontier model. On the first independent evaluation (Artificial Analysis Intelligence Index v4.3.x) it scores 30, versus 44–46 for the best Chinese open-weight models (Kimi K3, GLM-5.3, DeepSeek V4.1, MiMo-V2.6), 46 for xAI’s Grok 4.7, 52.6–53 for Google’s Gemini 4 Argon and OpenAI’s GPT-6 Astra, and 56–58 for Anthropic’s Claude Sonnet 5.5 and Opus 5.5. Reflection’s own launch table (parity with GLM-5.2 on reasoning) is therefore optimistic relative to independent measurement."),
  h2("1.2 What this means for building a competing model"),
  table(["Ambition level", "Independent target", "What it takes [E]", "Indicative 2-yr cost [E]"], [
    ["1. Beam parity", "AA Index ≈ 30; Beam KPIs in Section 11", "Adapt Beam (Apache 2.0) or train a ~500B-A23B MoE on ~24T tokens + large RL", "≈ US$10M (adapt) to ≈ US$175M (from scratch)"],
    ["2. Open-frontier parity", "AA ≈ 45 (Kimi/GLM/DeepSeek class)", "~1–2T total, 40–60B active, 30T+ tokens, RL compute ≥ pretraining", "≈ US$300–500M"],
    ["3. Closed-frontier challenger", "AA ≥ 53 (GPT-6 / Gemini 4 / Claude class)", "Multi-trillion-parameter models, 10^25–10^26 FLOPs, large proprietary data and RL environment fleets", "US$1B+ per generation; only realistic for hyperscaler-backed labs"],
  ], [1900, 2100, 3260, 2100]),
  gap(),
  h2("1.3 Recommendation"),
  num(runs(b("Do not attempt to out-compete Anthropic, OpenAI, Google or xAI head-on with a Beam-class model. "), "The independent gap is ~23–28 index points, and closed vendors hold ~89% of enterprise API spend.")),
  num(runs(b("Exploit the open-weight opportunity Beam has created. "), "Beam’s Apache 2.0 weights, reference implementation, Mirror CLI training harness and day-0 vLLM/SGLang support lower the entry cost for a specialised or sovereign model by one to two orders of magnitude.")),
  num(runs(b("Pursue a staged plan: "), "Stage 1 (0–6 months, ≈US$10M): adapt Beam for a chosen vertical/sovereign niche and prove the data, RL and evaluation stack. Stage 2 (6–18 months, gated): continued pre-training and full post-training (≈US$55M). Stage 3 (gated on traction and capital): a from-scratch next-generation model.")),
  num(runs(b("Compete on what closed vendors cannot offer: "), "on-premises/sovereign deployment, data residency, full customisation, predictable cost per token, and domain specialisation — the segments where open weights win despite a lower aggregate index.")),
  h2("1.4 Key numbers"),
  table(["Metric", "Value", "Basis"], [
    ["Beam total / active parameters", "501.3B / 23.4B", "[V] model card"],
    ["Experts / active per token", "256 / 8 (+3 dense layers)", "[V] model card, vLLM recipe"],
    ["Pretraining tokens / GPUs", "23.8T tokens / 6,144 GB300, < 4 weeks", "[V]"],
    ["RL compute", "10,500 GB300 × ~4 weeks; >100M rollouts", "[V]"],
    ["SWE-bench Verified (BF16)", "81.6%", "[V] model card"],
    ["Artificial Analysis Intelligence Index", "30 (19th of 120)", "[I]"],
    ["Leader on same index", "Claude Opus 5.5 (Max) = 58", "[I]"],
    ["Open-weight share of enterprise API spend", "11% (down from 19%)", "[S] Menlo Ventures, late 2025"],
    ["Estimated Beam training compute", "≈ 3.3 × 10^24 FLOPs pretraining; ≈ 11–14M GPU-hours in total", "[E]"],
  ], [3600, 3000, 2760]),
);

// =============== 2. Scope & method ===============
add(h1("2. Purpose, scope and methodology"),
  p("This document is the technical and commercial baseline for a programme to research, design and build a new large language model with capabilities equivalent to — and ideally exceeding — Reflection AI’s Beam. It is written for the founding technical team, investors and partners."),
  h2("2.1 Questions answered"),
  bullet("What exactly is Beam (architecture, data, training, release package) and what remains undisclosed?"),
  bullet("How does Beam perform against Anthropic, OpenAI, Google, xAI, Meta and the leading open-weight models — on vendor-reported and on independent benchmarks?"),
  bullet("What KPIs define “Beam-equivalent”, and what would it take to reach higher tiers?"),
  bullet("What are the strategic options, costs, team, infrastructure, risks and go-to-market paths for a new model?"),
  bullet("Which research directions should shape the design?"),
  h2("2.2 Sources and confidence"),
  p("Primary sources: Reflection’s launch and availability posts, the Hugging Face model cards (BF16, FP8, NVFP4), the Hugging Face, AWS, Google Cloud, NVIDIA and vLLM integration pages, and vendor system cards. Independent sources: Artificial Analysis, Vals AI and public leaderboards. Secondary sources: technical press and analyst blogs. Direct page retrieval was restricted in the research environment, so some figures were taken from search-indexed extracts of these sources; each is tagged with a confidence label (see document control)."),
  h2("2.3 Known limitations"),
  bullet("Beam’s full technical report is still marked “coming”; hidden size, optimiser and data mixture are not yet public."),
  bullet("Closed vendors publish different benchmark versions (e.g., Terminal-Bench 4.0 vs 2.1; SWE-bench Pro vs Verified). Only the Artificial Analysis index puts all models on one scale."),
  bullet("Market-size estimates vary widely by definition; they are used directionally only."),
);

// =============== 3. Reflection AI ===============
add(h1("3. Reflection AI: company profile"),
  kv([
    ["Founded", "2024, by former Google DeepMind researchers"],
    ["Leadership", "Misha Laskin (CEO); Ioannis Antonoglou (CTO, co-creator of AlphaGo)"],
    ["Mission", "An American open frontier lab producing open-weight models competitive with leading Chinese open models"],
    ["Funding", "US$2B (Oct 2025) led by NVIDIA (~US$800M) at US$8B valuation; reported talks in Mar 2026 for US$2.5B at ~US$25B pre-money; one aggregator lists US$27.5B (May 2026) [S]"],
    ["Products", "Asimov (code-research agent); Beam 1.0 (open weights); Mirror CLI (open-source agent harness used to train Beam)"],
    ["Ecosystem", "40+ launch partners incl. AWS, Google Cloud, Microsoft Azure, NVIDIA, Hugging Face, Baseten, Ollama, LM Studio"],
    ["Strategic news", "Reports (Oct 2026) of early NVIDIA talks to acquire, invest further in, or license technology from Reflection [S]"],
  ]),
  h2("3.1 Timeline"),
  table(["Date", "Event"], [
    ["2024", "Company founded"],
    ["Oct 2025", "US$2B round led by NVIDIA"],
    ["Mar–May 2026", "Reported raise at US$25–27.5B valuation [S]"],
    ["5 Oct 2026", "Beam announced: specs, launch benchmark table, waitlisted beta API (262K context)"],
    ["16 Oct 2026", "Weights released on Hugging Face (BF16, FP8, NVFP4) with tokenizer, reference implementation, DFlash2 drafter and Mirror CLI; day-0 support in Transformers, vLLM and SGLang; Beam Safety Technical Report published"],
    ["Oct 2026", "Artificial Analysis scores Beam 30 on Intelligence Index v4.3.x"],
    ["Pending", "Full Beam technical report"],
  ], [2000, 7360]),
  note("Not to be confused with “Reflection Llama-3.1 70B” (2024), an unrelated fine-tune by a different team."),
);

// =============== 4. Specification ===============
add(h1("4. Beam 1.0: complete specification"),
  table(["Attribute", "Value", "Source"], [
    ["Official name", "Beam-1.0-501B-A23B", "[V] Hugging Face"],
    ["Type", "Decoder-only sparse MoE transformer, text-only", "[V]"],
    ["Total / active parameters", "501.3B / 23.4B (≈4.7% active)", "[V] model card"],
    ["Layers", "52 = 3 dense-MLP layers + 49 MoE layers", "[V] model card"],
    ["Experts", "256 routed experts per MoE layer; top-8 routing", "[V] model card, vLLM recipe"],
    ["Attention", "Gated GQA; 512-token sliding-window attention on most layers; full-attention layers use NoPE (no positional encoding)", "[V] card; [S] vLLM recipe"],
    ["Full-attention layers", "9 of 52 keep a full-length KV cache", "[S] GGUF quantiser notes"],
    ["Normalisation / residual", "SandwichNorm; depth-based residual scaling; FP32 residual accumulation", "[V] launch post"],
    ["Load balancing", "Auxiliary-loss-free expert-bias balancing with cosine-decayed bias updates; busiest expert ≈1.04× mean", "[V] launch post"],
    ["Vocabulary", "152K tokens", "[V] model card"],
    ["Context length", "1M tokens (262K on the beta API)", "[V]"],
    ["Pretraining", "23.8T tokens; 6,144 GB300 GPUs; < 4 weeks", "[V]"],
    ["Data curation", "~95% of raw web tokens discarded", "[S] MarkTechPost"],
    ["RL post-training", "10,500 GB300 GPUs; ~4 weeks; >100M rollouts; ~1M environments; ~1.3B sandboxes; 256K max rollout context; Mirror CLI harness", "[V]"],
    ["Release formats", "BF16, FP8, NVFP4 checkpoints; DFlash2 speculative-decoding drafter (~1.93B params)", "[V]; drafter size [S]"],
    ["License", "Apache 2.0 (weights, tokenizer, reference implementation)", "[V]"],
    ["Hidden size, heads, optimiser, data mixture", "Not yet disclosed", "—"],
  ], [2400, 4560, 2400]),
);

// =============== 5. Architecture ===============
add(h1("5. Architecture deep-dive"),
  ...img("fig1_architecture.png", 6.5, "Figure 1 — Beam 1.0 architecture as confirmed by the model card and serving recipes."),
  h2("5.1 Sparse MoE backbone: 256 experts, top-8"),
  p("The first three layers use dense MLPs (a common choice for stabilising early representations); the remaining 49 layers replace the MLP with 256 fine-grained routed experts, of which a learned router selects 8 per token. With 8/256 experts active, each token touches ~3% of expert parameters; together with attention and embeddings this yields 23.4B active parameters. This configuration closely follows the DeepSeek-V3 lineage (256 routed experts, top-8), now the de-facto standard for open MoE models."),
  h2("5.2 Auxiliary-loss-free balancing with cosine-decayed bias updates"),
  p("Each expert has a bias added to its routing score for top-k selection only. Under-loaded experts have their bias raised after each step, over-loaded experts lowered. Reflection decays the bias-update step on a cosine schedule so routing settles late in training. Reported result: the busiest expert averages 1.04× the mean load — near-perfect utilisation, which directly improves expert-parallel throughput and avoids token dropping."),
  h2("5.3 Attention: 512-token sliding window + sparse NoPE global layers"),
  p("Most layers attend only to the previous 512 tokens, so their KV cache is constant regardless of context length. Only 9 of the 52 layers attend over the full sequence, and these use no positional encoding (NoPE), relying on the local layers to carry order information — a design that has been shown to extrapolate well to long contexts. Grouped-query attention (GQA) further shrinks KV size, and an elementwise output gate suppresses attention sinks."),
  table(["KV-cache illustration at 1M tokens [E]", "All 52 layers global", "Beam (9 global + 43 SWA-512)"], [
    ["Relative KV memory", "100%", "≈ 17% (9/52; SWA layers negligible)"],
    ["Relative attention FLOPs at long context", "100%", "≈ 17–20%"],
    ["Practical effect", "Long contexts need many GPUs for cache", "1M context feasible on a single 8-GPU node in FP8/NVFP4"],
  ], [3800, 2600, 2960]),
  h2("5.4 Stability components"),
  bullet(runs(b("SandwichNorm: "), "normalisation before and after each sub-layer to bound activation growth.")),
  bullet(runs(b("Gated attention output: "), "a learned sigmoid gate on attention outputs; reduces loss spikes and attention sinks.")),
  bullet(runs(b("Depth-based residual scaling: "), "scales residual-branch contributions by depth for stable signal propagation.")),
  bullet(runs(b("FP32 residual accumulation: "), "keeps the residual stream in FP32 while matrix multiplies run in low precision.")),
  h2("5.5 Inference accelerators"),
  p("Beam ships with DFlash2, a small (~1.93B) speculative-decoding drafter that proposes several tokens for the main model to verify in parallel, targeted at the latency of agentic loops. Reflection’s recipes recommend it for both vLLM and SGLang. Google Cloud reports day-0 serving with up to 5× performance improvement on its stack [V]."),
);

// =============== 6. Training ===============
add(h1("6. Training recipe"),
  ...img("fig2_pipeline.png", 6.5, "Figure 2 — Beam’s disclosed training pipeline."),
  table(["Phase", "Disclosed facts", "Implication for a replication"], [
    ["Data curation", "~95% of raw web discarded; 152K-vocab tokenizer", "Build model-based quality filters, deduplication, decontamination and synthetic code/maths before pre-training"],
    ["Pre-training", "23.8T tokens; 6,144 GB300; < 4 weeks", "≈ 3.3 × 10^24 FLOPs (6 × 23.4B × 23.8T) [E]; requires ~20–35% MoE MFU"],
    ["Mid-training", "Context extended to 1M tokens", "Long-document and repository-level data; staged sequence-length increase"],
    ["SFT", "Not detailed", "Seed tool-use formats, multi-turn agent traces and reasoning style"],
    ["RL", "10,500 GB300 × ~4 weeks; >100M rollouts; ~1M environments; ~1.3B sandboxes; Mirror CLI", "Asynchronous RL, sandbox orchestration and verifiable rewards are core engineering, not afterthoughts"],
  ], [1700, 3600, 4060]),
  h2("6.1 Mirror CLI: the harness Beam was trained against"),
  p("Reflection open-sourced Mirror CLI, the terminal-agent harness used during RL (reported as a Python front-end for prompts, skills and tooling with a Rust agent core). Training against a production harness means the model’s agentic behaviour matches how it will be deployed. For a new model, adopting Mirror CLI (or an equivalent) as both the RL environment interface and the evaluation harness removes a large source of train/serve mismatch."),
  h2("6.2 Compute split"),
  p("In GPU-hours Beam’s RL phase (≈7.1M GB300-hours) was roughly 2–3× its pre-training phase (≈2.5–3M) [E]. This inverts the historical ratio and reflects the 2026 consensus that agentic capability is mostly bought in post-training."),
);

// =============== 7. Release & deployment ===============
add(h1("7. Release package, deployment and hardware"),
  h2("7.1 Checkpoints and quality by precision (model card)"),
  table(["Benchmark", "BF16", "FP8", "NVFP4"], [
    ["SWE-bench Verified", "81.6", "80.2", "80.7"],
    ["SWE-bench Pro v1", "66.1", "65.9", "65.6"],
    ["Terminal-Bench 2.1", "80.1", "78.9", "78.4"],
    ["τ³-bench [S]", "36.1", "36.8", "35.4"],
  ], [3360, 2000, 2000, 2000]),
  note("Quantisation to FP8/NVFP4 costs ≤ 1.7 points on these agentic benchmarks — a strong argument for NVFP4-first serving."),
  h2("7.2 Memory footprint and minimum hardware [E]"),
  table(["Format", "Weights size (approx.)", "Minimum practical deployment"], [
    ["BF16", "≈ 1.0 TB (≈1,093 GB incl. overhead per one sizing tool)", "16× H200/B200 or 8× GB300/B300 (288 GB each)"],
    ["FP8", "≈ 0.5 TB", "8× H200 (141 GB) or 8× B200/B300"],
    ["NVFP4", "≈ 0.28 TB", "4–8× Blackwell-class GPUs"],
    ["Community GGUF (Q4)", "≈ 0.28–0.3 TB", "Large-memory workstations / CPU+GPU offload (unsloth, AtomicChat quantisations)"],
  ], [1800, 3600, 3960]),
  h2("7.3 Where Beam runs today"),
  bullet("Frameworks: Hugging Face Transformers, vLLM and SGLang (day-0), NVIDIA NeMo AutoModel and Dynamo recipes."),
  bullet("Clouds and platforms: AWS, Google Cloud, Microsoft Azure, NVIDIA, Baseten, and 40+ partners; local via Ollama and LM Studio."),
  bullet("Pricing: no Reflection rate card; partner pricing varies (Baseten listed at launch). Artificial Analysis had not yet measured speed or price."),
);

// =============== 8. Benchmarks ===============
add(h1("8. Performance analysis"),
  h2("8.1 Model scale in context"),
  ...img("fig3_params.png", 6.5, "Figure 3 — Total and active parameters of Beam and open-weight competitors."),
  h2("8.2 Vendor-reported: Beam vs. open-weight peers"),
  ...img("fig4_bench_open.png", 6.5, "Figure 4 — Reflection’s launch table (5 Oct). NR = not reported."),
  table(["Benchmark", "Beam", "GLM 5.2", "Kimi K3", "Qwen 3.8 Max"], [
    ["SWE-Bench Pro v1", "65.5 (66.1 card)", "62.1", "NR", "67.7"],
    ["SWE-Bench Pro v2-Hard", "77.2", "NR", "88.2", "NR"],
    ["SWE-Bench Verified", "80.9 (81.6 card)", "NR", "NR", "NR"],
    ["Terminal-Bench 2.1", "80.1", "81.0", "88.3", "86.6"],
    ["GPQA Diamond", "90.5 (90.6 card)", "91.2", "93.5", "92.6"],
    ["HLE (no tools)", "36.2", "40.5", "46.9", "43.6"],
    ["AIME 2026", "97.8", "99.2", "NR", "NR"],
  ], [2960, 1800, 1500, 1500, 1600]),
  h2("8.3 Independent: Artificial Analysis Intelligence Index"),
  ...img("fig8_aa_index.png", 6.5, "Figure 5 — The only single-scale, single-harness comparison across all vendors."),
  p("The independent index tells a different story from the vendor table. Beam (30) ranks 19th of 120 models — above the median for open models of its size and ahead of NVIDIA’s Nemotron 3.5 Super VL (26) — but 14–16 points behind the leading Chinese open-weight models and 23–28 points behind the closed frontier. Independent reviewers also note that some Reflection-reported agentic numbers (e.g., AutomationBench) were not reproduced. One clear strength: Beam used ~79M output tokens on the index suite versus a comparison median of ~130M, i.e. it reaches its answers with ~40% fewer tokens."),
  h2("8.4 Vendor-by-vendor comparison"),
  ...img("fig5_bench_frontier.png", 6.5, "Figure 6 — Best available published numbers on three common benchmarks."),
  table(["Vendor / model", "Weights", "AA Index [I]", "Headline published results", "List price in/out per 1M"], [
    ["Anthropic — Claude Opus 5.5 (Sep 2026)", "Closed", "58 (Max)", "SWE-bench Pro 89.9 [V]; Terminal-Bench 4.0 66.4 [V] (59.6–61.6 [I]); HLE 64.4 / 67.7 with tools [V]", "$4 / $20"],
    ["Anthropic — Claude Sonnet 5.5", "Closed", "56", "Terminal-Bench 4.0 70.6; SWE-bench Pro 81.3 [V]", "$2 / $10"],
    ["Anthropic — Claude Fable 5.1", "Closed", "≈53", "SWE-bench Pro 81.2; HLE 60.9; GPQA 92.6 [V/S]", "$10 / $50"],
    ["OpenAI — GPT-6 Astra", "Closed", "≈53", "Terminal-Bench 4.0 ~58–60 [I/S]; GPQA 96.1 [S]", "$10 / $50"],
    ["OpenAI — GPT-6.1 Sol", "Closed", "—", "Terminal-Bench Science 57.0; OSWorld 2.0 71.4 [S]", "$2 / $10"],
    ["Google — Gemini 4 Argon", "Closed", "52.6", "Roughly level with GPT-6 Astra on the index [I]", "—"],
    ["Google — Gemini 3.8 Flash", "Closed", "—", "GPQA 95.3–95.6; Terminal-Bench 2.1 89.4–90.8; SWE-bench Pro 61.6 [S]", "$0.75 / $3.75"],
    ["xAI — Grok 4.7", "Closed", "46", "Terminal-Bench 4.0 37.6 [V] vs 26 [S, independent]; GPQA (4.6) 94.9 [S]", "$2 / $6"],
    ["Meta — Muse Spark 1.3", "Closed", "61 on v4.1.1 (different index version)", "Terminal-Bench 2.1 88.8 [V]; Meta’s self-reports have exceeded verified scores before", "—"],
    ["Meta — Llama 4 Maverick", "Open (Llama licence)", "—", "GPQA 69.8 [V]; Meta’s latest open generation", "—"],
    ["Moonshot — Kimi K3", "Open", "44", "Terminal-Bench 2.1 88.3; GPQA 93.5; HLE 46.9 [V via Reflection]", "—"],
    ["Z.ai — GLM-5.3 (Max)", "Open", "45", "GLM-5.2: Terminal-Bench 2.1 81.0; GPQA 91.2 [V]", "—"],
    ["DeepSeek — V4.1 Flash", "Open", "45", "1.6T/49B (V4 Pro) [S]", "$0.14 / $0.28 (V4 Flash)"],
    ["Xiaomi — MiMo-V2.6-Pro", "Open", "46", "Top open model on the index [I]", "—"],
    ["Reflection — Beam 1.0", "Open (Apache 2.0)", "30", "SWE-bench Verified 81.6; Terminal-Bench 2.1 80.1; GPQA 90.6 [V]", "Partner-dependent"],
  ], [2300, 1200, 1300, 3260, 1300]),
  gap(),
  h3("Notes on specific vendors"),
  bullet(runs(b("xAI (Grok): "), "xAI (now operating as SpaceXAI per some reports) competes primarily on price ($2/$6) and speed; its agentic coding scores lag the leaders, and vendor vs independent Terminal-Bench numbers differ by ~12 points.")),
  bullet(runs(b("Meta: "), "Meta has split its line: the Muse series (Spark, Wonder) is closed and powers Meta AI; Llama 4 remains the latest open generation. A Beam-class open model competes with Llama, not Muse.")),
  bullet(runs(b("Groq (the inference company, distinct from xAI’s Grok): "), "does not train frontier models; it hosts open models on LPU hardware and is a potential deployment partner. NVIDIA licensed its technology in Dec 2025 (~US$20B).")),
  h2("8.5 Efficiency: capability per active parameter"),
  ...img("fig6_efficiency.png", 5.5, "Figure 7 — GPQA Diamond vs. active parameters (vendor-reported)."),
  p("On vendor numbers Beam reaches ~97% of Kimi K3’s GPQA score with under half its active parameters. Independent results narrow the efficiency story to serving economics: Beam’s 23.4B active parameters, sparse global attention, NVFP4 checkpoint and drafter make it among the cheapest models of its size to run per token."),
);

// =============== 9. Market ===============
add(h1("9. Market and competitive landscape"),
  h2("9.1 Market size and structure"),
  table(["Indicator", "Value", "Source / note"], [
    ["Enterprise generative-AI spend", "US$37B in 2025, up from US$11.5B (≈3.2×)", "[S] Menlo Ventures"],
    ["Global LLM market (narrow definition)", "US$10.6–12.8B in 2026", "[S] market-research aggregators; definitions vary"],
    ["Total AI spending (all categories)", "≈ US$2.5T in 2026 (+47% y/y)", "[S] Gartner via aggregators"],
    ["Enterprise API share — Anthropic / OpenAI", "40% / 27%", "[S] Menlo Ventures, late 2025"],
    ["U.S. businesses paying (Aug 2026)", "Anthropic 43.8%, OpenAI 39.8%", "[S] Ramp transaction data"],
    ["Open-weight share of enterprise workloads", "11–13%, down from 19%", "[S] Menlo Ventures"],
  ], [3300, 3000, 3060]),
  ...img("fig10_market_share.png", 6.0, "Figure 8 — Enterprise LLM API spend share."),
  p("Implication: enterprise demand is concentrated on the most capable closed models, and open-weight share has been falling. An open model must therefore win on dimensions that closed APIs cannot match — control, privacy, residency, customisation and cost — rather than on leaderboard rank alone."),
  h2("9.2 Price vs. capability"),
  ...img("fig9_price_vs_iq.png", 6.0, "Figure 9 — List output price vs. independent intelligence index."),
  table(["Model", "Input $/1M", "Output $/1M", "Positioning"], [
    ["Claude Fable 5.1 / GPT-6 Astra", "10", "50", "Premium frontier"],
    ["Claude Opus 5.5", "4", "20", "Highest independent index"],
    ["Claude Sonnet 5.5 / GPT-6.1 Sol", "2", "10", "Mid-price workhorse"],
    ["Grok 4.7", "2", "6", "Price-led challenger"],
    ["Gemini 3.8 Flash", "0.75", "3.75", "Fast/cheap tier"],
    ["Claude Haiku 5.5", "0.10", "0.50", "Lowest-cost production tier"],
    ["DeepSeek V4 Flash", "0.14", "0.28", "Open-weight price floor"],
  ], [3200, 1600, 1600, 2960]),
  note("Prices from third-party aggregators as of 2–17 Oct 2026; verify against vendor rate cards."),
  h2("9.3 Competitor profiles"),
  table(["Competitor", "Strengths", "Weaknesses / opening for a new entrant"], [
    ["Anthropic", "Top independent scores; leading enterprise share; strongest agentic coding", "Closed; premium pricing; no on-prem weights"],
    ["OpenAI", "Broadest consumer reach; strong reasoning; wide model range", "Closed; enterprise share below Anthropic"],
    ["Google", "Multimodal, long context, cheapest fast tier, cloud distribution", "Closed weights for Gemini; Gemma open models smaller"],
    ["xAI (Grok)", "Low prices, speed, X distribution", "Lags on agentic coding; vendor/independent score gaps"],
    ["Meta", "Massive distribution; Llama brand in open source", "Frontier Muse models closed; Llama 4 ageing"],
    ["Chinese open labs (DeepSeek, Moonshot, Z.ai, Alibaba, Xiaomi)", "Best open-weight scores; very low prices; rapid release cadence", "Geopolitical and procurement restrictions in US/EU government and regulated sectors"],
    ["Reflection (Beam)", "US-built, Apache 2.0, efficient, strong ecosystem launch", "Independent score well below Chinese open leaders; full report pending"],
    ["Mistral, NVIDIA (Nemotron)", "European / hardware-vendor open models", "Lower independent scores; narrower focus"],
  ], [2200, 3500, 3660]),
  h2("9.4 Where a new Beam-class model can win"),
  num(runs(b("Sovereign and government AI: "), "US/EU/allied governments that cannot procure Chinese open weights and need on-prem deployment.")),
  num(runs(b("Regulated industries: "), "finance, healthcare, defence and legal, where data residency and auditability outweigh a few index points.")),
  num(runs(b("Vertical specialists: "), "a model post-trained on proprietary domain environments (e.g., chip design, cybersecurity, scientific code) can beat general frontier models on narrow tasks.")),
  num(runs(b("Cost-sensitive agentic workloads at scale: "), "high-volume coding agents and background automation where per-token cost dominates.")),
  num(runs(b("Language and regional markets: "), "under-served languages (e.g., Hebrew/Arabic, regional European languages) where targeted data yields outsized gains.")),
);

// =============== 10. Gap analysis ===============
add(h1("10. Gap analysis: what it takes to compete"),
  table(["Tier", "Representative models", "AA Index", "Typical scale", "Gap from Beam"], [
    ["Closed frontier", "Claude Opus/Sonnet 5.5, GPT-6 Astra, Gemini 4 Argon", "53–58", "Undisclosed; likely multi-trillion parameters and 10^25–10^26 FLOPs", "+23 to +28"],
    ["Price-led closed", "Grok 4.7", "46", "Undisclosed", "+16"],
    ["Open frontier", "MiMo-V2.6-Pro, GLM-5.3, DeepSeek V4.1, Kimi K3", "44–46", "0.75–2.8T total; 40–95B active", "+14 to +16"],
    ["Beam class", "Beam 1.0", "30", "501B / 23B", "—"],
  ], [1600, 2900, 1100, 2400, 1360]),
  gap(),
  h2("10.1 Levers that close the gap (in order of expected return)"),
  num(runs(b("Post-training scale and environment diversity: "), "more RL compute, more verifiable environments, better reward design.")),
  num(runs(b("Active-parameter budget: "), "the open leaders use 40–95B active vs Beam’s 23B; doubling active parameters is the most direct capability lever, at proportionally higher serving cost.")),
  num(runs(b("Data quality and quantity: "), "30T+ curated tokens with larger synthetic reasoning/code shares.")),
  num(runs(b("Total capacity: "), "1–2T total parameters with the same active budget raise knowledge capacity without raising per-token compute.")),
  num(runs(b("Evaluation-driven iteration: "), "independent-harness evaluation each checkpoint to avoid optimising vendor-specific harnesses.")),
);

// =============== 11. KPIs ===============
add(h1("11. KPI framework"),
  p("“PKI” in the original brief is interpreted as KPI — key performance indicators. Three KPI levels are defined. All must be measured on a fixed, independent harness, and Beam itself (now downloadable) must be re-run on the same harness as the baseline."),
  table(["KPI", "Level 1: Beam parity", "Level 2: Open-frontier", "Level 3: Frontier challenger"], [
    ["AA Intelligence Index (or equivalent composite)", "≥ 30", "≥ 45", "≥ 53"],
    ["SWE-bench Verified", "≥ 81", "≥ 85", "≥ 90"],
    ["SWE-bench Pro (v1 / public)", "≥ 66", "≥ 72", "≥ 85"],
    ["Terminal-Bench 2.1", "≥ 80", "≥ 87", "≥ 90"],
    ["GPQA Diamond", "≥ 90", "≥ 93", "≥ 95"],
    ["HLE (no tools)", "≥ 36", "≥ 45", "≥ 60"],
    ["AIME (current year)", "≥ 97", "≥ 99", "≥ 99"],
    ["Active / total parameters", "≤ 25B / ≤ 550B", "≤ 50B / ≤ 1.5T", "Open"],
    ["Effective context (RULER ≥ 85%)", "≥ 256K (1M stretch)", "≥ 1M", "≥ 1M"],
    ["Output tokens per task vs. median", "≤ 0.65×", "≤ 0.65×", "≤ 0.75×"],
    ["Quality loss at NVFP4", "≤ 2 pts", "≤ 1.5 pts", "≤ 1 pt"],
    ["Expert load imbalance (max/mean)", "≤ 1.10×", "≤ 1.05×", "≤ 1.05×"],
    ["Serving cost per 1M output tokens (self-hosted) [E]", "≤ US$0.60", "≤ US$1.00", "≤ US$3.00"],
    ["Safety", "Pass published red-team + bio/cyber uplift evals", "Same + third-party audit", "Frontier safety framework"],
  ], [3000, 2100, 2100, 2160]),
);

// =============== 12. Strategy & cost ===============
add(h1("12. Strategic options and cost model"),
  ...img("fig11_options_cost.png", 6.0, "Figure 10 — Indicative two-year cost of three routes to a Beam-class model."),
  table(["", "A. Adapt Beam", "B. Continued pre-train + full post-train", "C. From-scratch Beam-class"], [
    ["What", "SFT + RL on Beam weights for a target domain/language; own evaluation & serving", "Add 3–6T domain tokens to Beam, extend context, run Beam-scale RL", "Own architecture, tokenizer, 24T+ token pre-train, full RL"],
    ["Time to first model", "3–6 months", "9–12 months", "15–18 months"],
    ["Compute [E]", "≈ 0.5–1M GPU-hrs (≈ US$3M)", "≈ 4–6M GPU-hrs (≈ US$18M)", "≈ 10–14M GPU-hrs (≈ US$55M)"],
    ["Team", "15–25", "40–60", "80–120"],
    ["Indicative 2-yr total", "≈ US$11M", "≈ US$56M", "≈ US$175M"],
    ["IP ownership", "Derivative of Apache 2.0 weights (permissive)", "Derivative, substantially modified", "Fully owned"],
    ["Ceiling", "Beam + domain gains", "Beam + 5–10 index points in domain", "Can scale to Level 2 with more capital"],
    ["Main risk", "Limited differentiation", "Catastrophic forgetting; data cost", "Execution risk; capital"],
  ], [1700, 2400, 2600, 2660]),
  gap(),
  h2("12.1 Compute model for option C [E]"),
  table(["Phase", "Estimate", "Assumptions"], [
    ["Pre-training", "≈ 3.3 × 10^24 FLOPs → ≈ 2–3M GB300-hours", "6 × 23.4B × 23.8T; 20–30% MFU on ~2.5 PFLOP/s BF16-dense (higher with FP8)"],
    ["Mid-training", "≈ 0.2–0.4M GPU-hours", "~1T tokens at increasing sequence length"],
    ["RL", "≈ 7M GPU-hours", "10,500 GPUs × 28 days (Beam’s disclosed scale)"],
    ["Ablations & pilots", "≈ 1–2M GPU-hours", "1–7B dense/MoE and ~60B-A3B pilots"],
    ["Inference for evals & synthetic data", "≈ 0.5–1M GPU-hours", "Data generation, LLM-judge, eval sweeps"],
    ["Total", "≈ 11–14M GPU-hours", "≈ US$40–70M at US$3–5 per GPU-hour"],
  ], [2400, 3200, 3760]),
  h2("12.2 Team composition for option C"),
  table(["Function", "Headcount", "Key roles"], [
    ["Pre-training & architecture", "15–20", "Scaling-law research, MoE kernels, distributed training"],
    ["Data", "15–25", "Crawling, filtering classifiers, synthetic data, licensing"],
    ["Post-training & RL", "15–25", "RL algorithms, reward modelling, environment builders"],
    ["Infrastructure", "15–20", "Cluster ops, storage, networking, sandbox fleet, inference"],
    ["Evaluation & safety", "8–12", "Benchmarks, red-teaming, policy"],
    ["Product, GTM, partnerships", "10–15", "Developer relations, enterprise sales, cloud partners"],
    ["Total", "≈ 80–120", "Fully loaded cost ≈ US$300–450K per head-year in the US [E]"],
  ], [2600, 1400, 5360]),
);

// =============== 13. Reference design ===============
add(h1("13. Reference architecture for the new model"),
  p("The proposed design adopts Beam’s confirmed choices as defaults and tests a small number of improvements suggested by 2026 research. All non-default choices must win in ablations at 1–7B scale and a ~60B-A3B pilot before full training."),
  table(["Component", "Default (Beam-matched)", "Candidate improvement to ablate", "Rationale"], [
    ["Size", "≈ 500B total / ≈ 23B active, 52 layers", "≈ 600B / 30B active", "Most direct capability lever; still fits 8-GPU node in NVFP4"],
    ["Dense layers", "First 3 layers dense", "1–4 dense", "Early-layer stability"],
    ["Experts", "256 routed, top-8", "+1 shared expert; 384 routed", "DeepSeek-V3/V4 practice"],
    ["Routing", "Aux-loss-free bias balancing, cosine-decayed updates", "Sequence-level balance regulariser", "Reproduce 1.04× load"],
    ["Attention", "Gated GQA; SWA-512 on 43 layers; 9 NoPE global layers", "Compressed/sparse global attention (DeepSeek-V4 style); MLA on global layers", "Further KV reduction at 1M context"],
    ["Norm / residual", "SandwichNorm; depth-scaled residual; FP32 residual accumulation", "+ QK-norm", "Low-precision stability"],
    ["Tokenizer", "≈ 152K BPE", "≈ 200K with target-language coverage", "Compression on code and target languages"],
    ["Prediction head", "Next-token", "Multi-token prediction (1–2 heads)", "Denser signal; built-in speculative decoding"],
    ["Optimiser", "AdamW (assumed)", "Muon for matrix params (per-expert), AdamW for others", "~2× token efficiency reported"],
    ["Precision", "BF16/FP8 training", "NVFP4 training on GB300", "≈ 2× matmul throughput"],
    ["Drafter", "DFlash2-style ~2B drafter", "MTP-head self-drafting", "Latency for agent loops"],
  ], [1500, 2700, 2700, 2460]),
);

// =============== 14. Data ===============
add(h1("14. Data strategy"),
  h2("14.1 Pre-training corpus target (≈ 24–30T tokens) [E]"),
  table(["Source", "Share", "Notes"], [
    ["Filtered web text", "40–50%", "Model-based quality classifiers; keep top ~5–10% (Beam discarded ~95%)"],
    ["Code (permissive licences)", "20–25%", "Repository-level packing; dedup against benchmarks"],
    ["Mathematics & science", "8–12%", "Papers, textbooks, formal proofs"],
    ["Synthetic reasoning / code", "10–15%", "Execution-verified solutions, rephrased high-quality text"],
    ["Multilingual & target languages", "5–10%", "Up-weight target markets (e.g., Hebrew, Arabic, EU languages)"],
    ["Books, reference, licensed data", "3–5%", "Licensed where required"],
  ], [3000, 1400, 4960]),
  h2("14.2 Data engineering requirements"),
  bullet("Exact and fuzzy deduplication (MinHash/LSH) at document and repository level."),
  bullet("Benchmark decontamination against every KPI benchmark, with n-gram and embedding checks."),
  bullet("PII removal, toxicity filtering, licence tracking and opt-out handling (copyright and EU AI Act transparency)."),
  bullet("Data-mixture ablations on 1–7B proxies; annealing-phase mixture with the highest-quality data."),
  bullet("Long-context corpus: repositories, books and multi-document sets for mid-training to 1M tokens."),
);

// =============== 15. Post-training ===============
add(h1("15. Post-training and agentic RL"),
  num(runs(b("SFT: "), "high-quality instruction, tool-use and agent-trace data in the target harness format (Mirror CLI-compatible).")),
  num(runs(b("Environment fleet: "), "target 10^5–10^6 verifiable environments — unit-test-graded repositories, terminal tasks, maths/science checkers, web and tool tasks — with sandbox orchestration able to run billions of executions (Beam used ~1.3B sandboxes).")),
  num(runs(b("RL algorithm: "), "GRPO-family policy optimisation with asynchronous rollout generation, importance-ratio clipping for off-policy correction, and 128K–256K rollout contexts.")),
  num(runs(b("Reward design: "), "prefer verifiable rewards; use reward models only where verification is impossible; guard against reward hacking with held-out graders.")),
  num(runs(b("Token efficiency: "), "add length-aware rewards to keep output tokens per task ≤ 0.65× the market median (Beam’s distinctive strength).")),
  num(runs(b("Safety training: "), "refusal calibration, cyber/bio uplift mitigations and system-prompt robustness before release.")),
);

// =============== 16. Infrastructure ===============
add(h1("16. Infrastructure plan"),
  table(["Component", "Requirement (option C) [E]", "Notes"], [
    ["Training cluster", "≈ 6,000 GB300-class GPUs for ~1 month (pre-train); ≈ 10,000 for ~1 month (RL)", "Rent from neoclouds/hyperscalers; reserve 9–12 months ahead"],
    ["Interconnect", "NVLink 72-GPU domains; 800G InfiniBand/Ethernet scale-out", "Expert parallelism is all-to-all heavy"],
    ["Storage", "≈ 10–20 PB for raw and processed data; fast checkpoint store", "Checkpoints ≈ 6 TB each with optimiser state"],
    ["Sandbox fleet", "CPU clusters / containers for RL environments (millions of concurrent sandboxes at peak)", "Often the hidden bottleneck"],
    ["Inference for RL", "Dedicated vLLM/SGLang rollout servers with weight sync", "Asynchronous actor–learner design"],
    ["Serving", "8-GPU nodes in NVFP4 with drafter; autoscaling", "Partner with clouds / Groq-class inference providers"],
  ], [2000, 4300, 3060]),
);

// =============== 17. Evaluation ===============
add(h1("17. Evaluation plan"),
  bullet("Fixed internal harness covering all Section 11 KPIs, run on every checkpoint with confidence intervals."),
  bullet("Re-run Beam 1.0, GLM-5.x, Kimi K3 and DeepSeek V4.x on the same harness; never compare against vendor tables alone."),
  bullet("Prefer unsaturated and contamination-resistant suites: SWE-bench Pro, Terminal-Bench 2.x/4.0, HLE, current-year AIME, RULER, τ-bench/τ³-bench, OSWorld."),
  bullet("Submit to Artificial Analysis, Vals AI and LMArena before launch for independent numbers."),
  bullet("Track token efficiency and cost per solved task alongside accuracy."),
  bullet("Domain-specific private evals for the chosen vertical; these define commercial success more than public indices."),
);

// =============== 18. Governance ===============
add(h1("18. Safety, legal and governance"),
  table(["Topic", "Consideration"], [
    ["Licence of base weights", "Apache 2.0 permits commercial use, modification and redistribution; preserve NOTICE and attribution; includes a patent grant"],
    ["EU AI Act", "General-purpose AI model obligations (technical documentation, copyright policy, training-data summary). Models above 10^25 training FLOPs are presumed to carry systemic risk — a Beam-class model (≈3–4 × 10^24 [E]) sits below this threshold; a Level 2/3 model likely exceeds it"],
    ["Copyright and data", "Licence tracking, opt-out compliance, provenance records; litigation risk is concentrated in training-data choices"],
    ["Export controls", "Track US rules on model weights and advanced compute; open-weight release decisions need legal review"],
    ["Safety evaluation", "Pre-release red-teaming, cyber and bio uplift evaluation, and a published safety report (Reflection published a Beam Safety Technical Report)"],
    ["Responsible release", "Staged release (API → weights), model card, acceptable-use policy"],
  ], [2400, 6960]),
);

// =============== 19. GTM ===============
add(h1("19. Go-to-market and business model"),
  table(["Model", "Description", "Revenue logic"], [
    ["Open weights + managed API", "Release weights; sell a hosted API and enterprise SLA", "Usage-based; low price per token"],
    ["Sovereign / on-prem licensing", "Packaged deployment for governments and regulated enterprises", "Annual licence + support; high contract value"],
    ["Vertical model", "Domain-specialised model and agent product", "Subscription per seat or per task"],
    ["Cloud marketplace", "List on AWS, Azure, Google Cloud, NVIDIA NIM", "Revenue share; distribution leverage"],
    ["Custom post-training services", "Fine-tune/RL on customer environments", "Project fees + recurring hosting"],
  ], [2400, 4000, 2960]),
  gap(),
  p("Beam’s launch is the template: an open-weights release integrated with 40+ partners on day 0, plus an open-source harness that makes developers productive immediately. A new entrant should secure cloud and inference partners before release and ship serving recipes for vLLM and SGLang on day 0."),
);

// =============== 20. Risks ===============
add(h1("20. Risks and mitigations"),
  table(["Risk", "Likelihood", "Impact", "Mitigation"], [
    ["Capability gap to frontier widens faster than the plan can close it", "High", "High", "Do not compete head-on; target niches where open weights win"],
    ["Chinese open models keep a 15-point lead at near-zero price", "High", "Medium", "Differentiate on provenance, sovereignty, support and domain quality"],
    ["Vendor benchmark claims do not replicate", "Medium", "High", "Independent harness from day one; publish third-party results"],
    ["GPU availability / price shocks", "Medium", "High", "Multi-provider reservations; NVFP4 to cut compute needs"],
    ["RL infrastructure underestimated", "High", "High", "Start environment fleet and sandbox orchestration in month 1"],
    ["Key-person dependence / hiring", "Medium", "High", "Competitive equity; partner with universities"],
    ["Legal (copyright, export control, AI Act)", "Medium", "Medium–High", "Provenance tracking; legal review gates"],
    ["Consolidation (e.g., Reflection acquired; Beam licence terms change for future versions)", "Medium", "Medium", "Mirror Apache 2.0 Beam 1.0 weights now; the grant is irrevocable for released versions"],
    ["Commoditisation of open models", "High", "Medium", "Monetise services, deployment and vertical products rather than raw weights"],
  ], [3200, 1200, 1200, 3760]),
);

// =============== 21. Roadmap ===============
add(h1("21. Roadmap and decision gates"),
  ...img("fig7_roadmap.png", 6.5, "Figure 11 — Programme roadmap for option C; options A and B use phases 0–2 and 5–7."),
  table(["Gate", "Timing", "Go criteria"], [
    ["G0 — Baseline", "Month 1", "Beam 1.0 and peers reproduced on internal harness within ±2 pts of published values"],
    ["G1 — Adapt", "Month 6", "Adapted Beam beats base Beam by ≥ 5 pts on domain evals; first paying pilot customer"],
    ["G2 — Scale decision", "Month 9", "Pilot (60B-A3B) matches scaling-law predictions; data engine ≥ 20T curated tokens; funding ≥ US$150M secured"],
    ["G3 — Full pre-train", "Month 11", "Loss and ablation targets met; no instability in first 10% of tokens"],
    ["G4 — Release", "Month 17", "Level 1 KPIs met independently; safety report complete; ≥ 10 launch partners"],
  ], [2200, 1400, 5760]),
);

// =============== 22. Recommendations ===============
add(h1("22. Recommendations from current research"),
  h2("22.1 Architecture"),
  num(runs(b("Fine-grained MoE with aux-loss-free balancing "), "(256 experts, top-8) — confirmed in Beam and DeepSeek-V3/V4.")),
  num(runs(b("Aggressively sparse global attention: "), "Beam shows 9/52 global layers with SWA-512 elsewhere is viable at 1M context; evaluate DeepSeek-V4-style compressed sparse attention for the global layers.")),
  num(runs(b("NoPE on global layers "), "for length generalisation; RoPE (or none) on local layers.")),
  num(runs(b("Gated attention, SandwichNorm, QK-norm and FP32 residuals "), "for low-precision stability.")),
  num(runs(b("Multi-token prediction "), "for sample efficiency and self-speculative decoding (DeepSeek-V3/V4, MiMo-V2.6).")),
  h2("22.2 Optimisation and precision"),
  num(runs(b("Muon optimiser "), "for matrix parameters (Moonlight: ~2× compute efficiency vs AdamW; adopted by DeepSeek-V4 with per-expert orthogonalisation).")),
  num(runs(b("FP8 → NVFP4 training and serving: "), "Beam’s NVFP4 checkpoint loses ≤ 1.7 points; plan NVFP4 serving from day one.")),
  num(runs(b("WSD learning-rate schedules "), "to branch mid-training and annealing runs from a shared trunk.")),
  h2("22.3 Data and post-training"),
  num(runs(b("Data engine before model: "), "aggressive quality filtering and execution-verified synthetic data.")),
  num(runs(b("RL compute ≥ pre-training compute, "), "with a million-scale environment fleet and an open harness (Mirror CLI) shared between training and deployment.")),
  num(runs(b("Reward token efficiency "), "explicitly to reproduce Beam’s ~40% token saving.")),
  h2("22.4 Strategy"),
  num(runs(b("Start from Beam 1.0 weights. "), "Apache 2.0 makes it the cheapest credible starting point for a US/EU-provenance model.")),
  num(runs(b("Measure independently, always. "), "Beam’s gap between vendor and independent numbers is the central lesson of its launch.")),
  num(runs(b("Pick one axis to lead on "), "— sovereignty, a vertical, a language, or cost — rather than the general leaderboard.")),
);

// =============== 23. Downloads ===============
add(h1("23. Download links and access"),
  p(runs(b("Beam 1.0 is now downloadable "), "(released 16 October 2026, Apache 2.0). Verify the publisher is Reflection’s official Hugging Face organisation (“reflection”).")),
  table(["Resource", "Location", "Notes"], [
    ["Beam 1.0 (BF16)", "huggingface.co/reflection/Beam-1.0-501B-A23B", "≈ 1 TB"],
    ["Beam 1.0 (FP8)", "huggingface.co/reflection/Beam-1.0-501B-A23B-FP8", "≈ 0.5 TB"],
    ["Beam 1.0 (NVFP4)", "huggingface.co/reflection/Beam-1.0-501B-A23B-NVFP4", "≈ 0.28 TB; Blackwell"],
    ["Community GGUF", "huggingface.co/unsloth/Beam-1.0-501B-A23B-GGUF", "Unofficial quantisations"],
    ["Hugging Face launch blog", "huggingface.co/blog/reflection-beam", "Usage guide"],
    ["vLLM serving recipe", "recipes.vllm.ai/reflection/Beam-1.0-501B-A23B", "B300 / H200 commands"],
    ["NVIDIA Dynamo recipe", "docs.nvidia.com/dynamo/dev/recipes/beam-1-0-501b", "Disaggregated serving"],
    ["AWS guide", "aws.amazon.com/blogs/machine-learning/…reflection-beam…", "Agentic coding on AWS"],
    ["Google Cloud guide", "cloud.google.com/blog/…reflection-beam…", "Day-0, 5× performance claim"],
    ["Mirror CLI, DFlash2 drafter, Safety Technical Report", "Linked from Reflection’s “Beam is now available” post", "Open-source harness and drafter"],
    ["Open-weight peers", "huggingface.co/deepseek-ai · /moonshotai · /zai-org · /Qwen · /mistralai · /meta-llama · /openai", "Baselines for comparison"],
    ["Closed APIs", "console.anthropic.com · platform.openai.com · aistudio.google.com · console.x.ai", "API only"],
    ["Inference providers", "console.groq.com and others", "Hosted open models"],
  ], [2700, 4300, 2360]),
  gap(),
  p("Clickable links:"),
  ...[
    ["Beam 1.0 BF16 on Hugging Face", "https://huggingface.co/reflection/Beam-1.0-501B-A23B"],
    ["Beam 1.0 FP8 on Hugging Face", "https://huggingface.co/reflection/Beam-1.0-501B-A23B-FP8"],
    ["Beam 1.0 NVFP4 on Hugging Face", "https://huggingface.co/reflection/Beam-1.0-501B-A23B-NVFP4"],
    ["Unsloth GGUF quantisations", "https://huggingface.co/unsloth/Beam-1.0-501B-A23B-GGUF"],
    ["Hugging Face blog: Welcome Beam", "https://huggingface.co/blog/reflection-beam"],
    ["Reflection: Beam is now available", "https://reflection.ai/blog/beam-is-available"],
    ["vLLM recipe", "https://recipes.vllm.ai/reflection/Beam-1.0-501B-A23B"],
    ["NVIDIA Dynamo recipe", "https://docs.nvidia.com/dynamo/dev/recipes/beam-1-0-501b"],
    ["AWS: Reflection Beam for agentic coding", "https://aws.amazon.com/blogs/machine-learning/introducing-reflection-beam-an-open-weight-model-for-agentic-coding-on-aws/"],
    ["Google Cloud: day-0 support for Beam", "https://cloud.google.com/blog/topics/ai-infrastructure/day-0-support-with-5x-perf-for-reflection-beam-on-google-cloud"],
    ["Artificial Analysis: Beam article", "https://artificialanalysis.ai/articles/reflection-beam"],
    ["DeepSeek on Hugging Face", "https://huggingface.co/deepseek-ai"],
    ["Moonshot AI (Kimi) on Hugging Face", "https://huggingface.co/moonshotai"],
    ["Z.ai (GLM) on Hugging Face", "https://huggingface.co/zai-org"],
    ["Qwen on Hugging Face", "https://huggingface.co/Qwen"],
    ["Meta Llama on Hugging Face", "https://huggingface.co/meta-llama"],
  ].map(([t, u]) => bullet([link(t, u)])),
);

// =============== 24. Sources ===============
const sources = [
  ["Reflection — Introducing Beam", "https://reflection.ai/blog/introducing-beam"],
  ["Reflection — Beam is now available", "https://reflection.ai/blog/beam-is-available"],
  ["Hugging Face — reflection/Beam-1.0-501B-A23B model card", "https://huggingface.co/reflection/Beam-1.0-501B-A23B"],
  ["Hugging Face — Beam FP8 model card", "https://huggingface.co/reflection/Beam-1.0-501B-A23B-FP8"],
  ["Hugging Face — Welcome Beam (blog)", "https://huggingface.co/blog/reflection-beam"],
  ["vLLM — Beam serving recipe", "https://recipes.vllm.ai/reflection/Beam-1.0-501B-A23B"],
  ["NVIDIA NeMo AutoModel — Beam coverage", "https://docs.nvidia.com/nemo/automodel/model-coverage/large-language-models/reflection/Beam-1.0-501B-A23B"],
  ["Google Cloud — Day-0 support for Reflection Beam", "https://cloud.google.com/blog/topics/ai-infrastructure/day-0-support-with-5x-perf-for-reflection-beam-on-google-cloud"],
  ["AWS — Introducing Reflection Beam", "https://aws.amazon.com/blogs/machine-learning/introducing-reflection-beam-an-open-weight-model-for-agentic-coding-on-aws/"],
  ["Artificial Analysis — Reflection Beam", "https://artificialanalysis.ai/articles/reflection-beam"],
  ["Artificial Analysis — Intelligence Index v4.3.2", "https://artificialanalysis.ai/evaluations/artificial-analysis-intelligence-index"],
  ["OrcaRouter — Beam 1.0 scored 30 independently", "https://www.orcarouter.ai/blog/reflection-beam-1-0-open-weights-available"],
  ["Shattered.io — Beam at 13 days", "https://shattered.io/beam-first-13-days-reflection-501b-bet-tested-2026/"],
  ["TechCityAuthority — Beam second among US open models", "https://www.techcityauthority.com/2026/10/reflection-beam-open-weights-benchmarks.html"],
  ["Unite.AI — Beam open weights under Apache 2.0", "https://www.unite.ai/reflection-releases-beams-open-weights-under-apache-2-0-license/"],
  ["TechCrunch — Reflection debuts Beam", "https://techcrunch.com/2026/10/05/reflection-debuts-beam-a-open-weight-ai-model-to-rival-chinese-models-at-lower-compute-cost/"],
  ["MarkTechPost — Beam 501B open-weight MoE", "https://www.marktechpost.com/2026/10/05/reflection-ai-introduces-beam-a-501b-open-weight-moe-model-with-23b-active-parameters-for-coding-and-agentic-workloads/"],
  ["Kingy AI — Beam specs and benchmarks", "https://kingy.ai/news/reflection-beam-501b-open-weight-ai-guide/"],
  ["RuntimeWire — Beam benchmark scores", "https://runtimewire.com/article/reflection-ai-beam-open-weight-model"],
  ["Spheron — Beam VRAM requirements", "https://www.spheron.network/tools/gpu-recommender/reflection/Beam-1.0-501B-A23B/"],
  ["BenchLM — Beam 1.0 benchmarks", "https://benchlm.ai/models/reflection-beam"],
  ["Wikipedia — Reflection AI", "https://en.wikipedia.org/wiki/Reflection_AI"],
  ["Anthropic — Claude Fable 5.1 & Mythos 5.1 System Card", "https://www-cdn.anthropic.com/0339e6a7c5c7b87f5c07798616dc32c215d14235/Claude%20Fable%205.1%20&%20Claude%20Mythos%205.1%20System%20Card.pdf"],
  ["llm-stats — Claude Opus 5.5 launch", "https://llm-stats.com/blog/research/claude-opus-5-5-launch"],
  ["Digital Applied — Claude Opus 5.5 pricing and benchmarks", "https://www.digitalapplied.com/blog/claude-opus-5-5-launch-pricing-benchmarks-2026"],
  ["Vellum — GPT-6.1 Sol benchmarks", "https://www.vellum.ai/blog/gpt-6-1-sol-benchmarks-explained"],
  ["The Decoder — Grok 4.7 launch", "https://the-decoder.com/xai-launches-grok-4-7-at-bargain-prices-but-benchmarks-reveal-a-wide-gap-to-claude-and-gpt-6/"],
  ["Artificial Analysis — Muse Spark", "https://artificialanalysis.ai/articles/muse-spark-everything-you-need-to-know"],
  ["Interconnects — Meta shakes up the frontiers", "https://www.interconnects.ai/p/meta-shakes-up-the-frontiers-open"],
  ["DataCamp — Gemini 3.8 Flash", "https://www.datacamp.com/blog/gemini-3-8-flash-cyber"],
  ["pricepertoken — GPQA leaderboard", "https://pricepertoken.com/leaderboards/benchmark/gpqa"],
  ["BenchLM — LLM API pricing (Oct 2026)", "https://benchlm.ai/llm-pricing"],
  ["ComputingForGeeks — Open-source LLM comparison", "https://computingforgeeks.com/open-source-llm-comparison/"],
  ["DeepInfra — Kimi K3 vs DeepSeek V4 Pro vs GLM-5.2", "https://deepinfra.com/blog/kimi-k3-vs-deepseek-v4-pro-vs-glm-5-2"],
  ["Hostinger — LLM statistics 2026", "https://www.hostinger.com/tutorials/llm-statistics/"],
  ["Konabayev — LLM market share 2026", "https://konabayev.com/blog/llm-market-share/"],
  ["Voiceflow — Groq in 2026", "https://www.voiceflow.com/blog/groq"],
  ["DeepSeek-V4 technical report", "https://arxiv.org/html/2606.19348v1"],
  ["MiMo-V2.6: Scaling RL Towards Self-Improvement", "https://arxiv.org/html/2610.11959v1"],
  ["Muon is Scalable for LLM Training", "https://arxiv.org/pdf/2502.16982"],
  ["SOAP, Muon, and Beyond", "https://arxiv.org/pdf/2607.20548"],
  ["SWE-Bench Pro Verified", "https://arxiv.org/pdf/2609.08149"],
];
add(h1("24. Sources"),
  ...sources.map(([t, u]) => new Paragraph({ numbering: { reference: "src", level: 0 }, spacing: { after: 40 }, children: [link(t, u)] })),
);

// =============== Appendices ===============
add(h1("Appendix A. Formulas and worked estimates"),
  table(["Quantity", "Formula", "Beam example [E]"], [
    ["Training FLOPs", "≈ 6 × N_active × D_tokens", "6 × 23.4e9 × 23.8e12 ≈ 3.3 × 10^24"],
    ["GPU-hours", "FLOPs ÷ (peak FLOP/s × MFU × 3,600)", "3.3e24 ÷ (2.5e15 × 0.25 × 3,600) ≈ 1.5M (BF16); range 1.5–3M"],
    ["Decode FLOPs per token", "≈ 2 × N_active", "≈ 47 GFLOPs per output token"],
    ["Weight memory", "N_total × bytes per parameter", "501B × 2 B ≈ 1.0 TB (BF16); × 1 B ≈ 0.5 TB (FP8); × ~0.56 B ≈ 0.28 TB (NVFP4)"],
    ["KV cache share at long context", "global layers ÷ total layers", "9 ÷ 52 ≈ 17% of an all-global design"],
    ["Active ratio", "N_active ÷ N_total", "23.4 ÷ 501.3 ≈ 4.7%"],
  ], [2400, 3200, 3760]),
);
add(h1("Appendix B. Glossary"),
  table(["Term", "Meaning"], [
    ["MoE", "Mixture-of-Experts: many expert sub-networks per layer, few active per token"],
    ["501B-A23B", "501B total parameters, ~23B active per token"],
    ["Top-8 of 256", "Router selects 8 of 256 experts for each token"],
    ["SWA", "Sliding-window attention: attends only to the last N tokens (512 in Beam)"],
    ["NoPE", "No positional encoding on a layer"],
    ["GQA / MLA", "Grouped-query / multi-head latent attention: KV-cache reduction techniques"],
    ["SandwichNorm", "Normalisation before and after each sub-layer"],
    ["NVFP4", "NVIDIA 4-bit floating-point format on Blackwell GPUs"],
    ["Speculative decoding / drafter", "A small model proposes tokens that the large model verifies in parallel"],
    ["MTP", "Multi-token prediction heads"],
    ["Muon", "Optimiser that orthogonalises momentum updates for matrix parameters"],
    ["MFU", "Model FLOPs utilisation"],
    ["AA Intelligence Index", "Artificial Analysis composite of reasoning, knowledge, coding and agentic evaluations"],
    ["GPQA / HLE / AIME", "Graduate science QA / Humanity’s Last Exam / maths olympiad qualifier"],
    ["SWE-bench / Terminal-Bench / τ-bench", "Software-engineering, terminal-agent and tool-use agent benchmarks"],
    ["KPI", "Key performance indicator"],
  ], [2600, 6760]),
);

const doc = new Document({
  creator: "Beam research baseline", title: "Building a Beam-Class LLM — Research Baseline v2.0",
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
      children: [new TextRun({ text: "Building a Beam-Class LLM — Research Baseline v2.0", size: 16, color: INK2 })] })] }) },
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
