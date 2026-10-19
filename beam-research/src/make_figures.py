"""Generate charts and diagrams for the Beam research baseline document."""
import sys
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch

OUT = sys.argv[1]
BLUE, ORANGE, AQUA, YELLOW, VIOLET = "#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#6250d6"
GRAY, INK, INK2, GRID, SURF = "#b9b8b2", "#0b0b0b", "#52514e", "#e6e5e0", "#fcfcfb"
plt.rcParams.update({
    "font.family": "DejaVu Sans", "font.size": 10, "axes.edgecolor": GRID,
    "axes.labelcolor": INK2, "xtick.color": INK2, "ytick.color": INK2,
    "axes.spines.top": False, "axes.spines.right": False, "figure.facecolor": SURF,
    "axes.facecolor": SURF, "savefig.dpi": 200, "savefig.bbox": "tight",
})

def title(ax, t, sub=None):
    ax.set_title(t, loc="left", fontsize=13, fontweight="bold", color=INK, pad=22 if sub else 10)
    if sub:
        ax.text(0, 1.02, sub, transform=ax.transAxes, fontsize=9, color=INK2)

def box(ax, x, y, w, h, text, fc, ec=None, fs=9, color=INK, bold=False):
    ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.02,rounding_size=0.08",
                                fc=fc, ec=ec or fc, lw=1.2))
    ax.text(x + w / 2, y + h / 2, text, ha="center", va="center", fontsize=fs, color=color,
            fontweight="bold" if bold else "normal", wrap=True)

def arrow(ax, x1, y1, x2, y2, color=INK2):
    ax.add_patch(FancyArrowPatch((x1, y1), (x2, y2), arrowstyle="-|>", mutation_scale=12,
                                 color=color, lw=1.3))

# 1. Architecture overview -------------------------------------------------
fig, ax = plt.subplots(figsize=(11, 7.2)); ax.set_xlim(0, 11); ax.set_ylim(0, 7.2); ax.axis("off")
ax.text(0.1, 6.95, "Beam-1.0-501B-A23B: architecture (per Hugging Face model card, 16 Oct 2026)",
        fontsize=13, fontweight="bold", color=INK)
ax.text(0.1, 6.65, "501.3B total / 23.4B active · 52 layers (3 dense + 49 MoE) · 256 routed experts, top-8 · 152K vocabulary · 1M context\n"
        "Gated GQA + 512-token sliding-window attention; full-attention (NoPE) KV cache in only 9 of 52 layers",
        fontsize=8.5, color=INK2, va="top")
box(ax, 0.3, 0.4, 2.6, 0.55, "Token embedding (152K vocab)", "#dbe9f9", fs=9)
box(ax, 0.3, 5.15, 2.6, 0.55, "Final norm + LM head", "#dbe9f9", fs=9)
ax.add_patch(FancyBboxPatch((0.3, 1.2), 2.6, 3.7, boxstyle="round,pad=0.02,rounding_size=0.1", fc="#f1f0ec", ec=GRAY, lw=1))
ax.text(1.6, 4.65, "52 decoder layers", ha="center", fontsize=9.5, fontweight="bold", color=INK)
rows = [("D", "SWA attn + dense MLP (layers 0–2)")] + [("L", "SWA (512) attn + MoE")] * 4 + [("G", "Full attn (NoPE) + MoE")] + \
       [("L", "SWA (512) attn + MoE")] * 5 + [("G", "Full attn (NoPE) + MoE")]
for i, (k, t) in enumerate(rows):
    y = 1.35 + i * 0.27
    fc = {"D": "#fde3d6", "L": "#cfe0f5", "G": BLUE}[k]
    box(ax, 0.55, y, 2.1, 0.21, t, fc, fs=6.5, color="white" if k == "G" else INK)
ax.text(1.6, 1.24, "9 full-attention layers in total (spacing illustrative)", ha="center", fontsize=6.5, color=INK2)
arrow(ax, 1.6, 0.95, 1.6, 1.2); arrow(ax, 1.6, 4.9, 1.6, 5.15)
ax.add_patch(FancyBboxPatch((3.6, 0.35), 7.1, 5.4, boxstyle="round,pad=0.02,rounding_size=0.1", fc="white", ec=GRAY, lw=1))
ax.text(3.8, 5.5, "One Beam MoE decoder layer", fontsize=10.5, fontweight="bold", color=INK)
arrow(ax, 2.95, 3.0, 3.6, 3.0, color=GRAY)
x0 = 4.0
box(ax, x0, 0.6, 2.4, 0.45, "residual stream in (FP32 accum.)", "#eeeeea", fs=8)
box(ax, x0, 1.3, 2.4, 0.42, "Pre-Norm (SandwichNorm)", "#fde3d6", fs=8)
box(ax, x0, 1.95, 2.4, 0.6, "Gated GQA self-attention\nSWA-512 or full (NoPE)", "#cfe0f5", fs=8)
box(ax, x0, 2.8, 2.4, 0.42, "Elementwise output gate  σ(xW_g)", "#d4f0e5", fs=8)
box(ax, x0, 3.45, 2.4, 0.42, "Post-Norm (SandwichNorm)", "#fde3d6", fs=8)
box(ax, x0, 4.1, 2.4, 0.5, "+ residual, depth-scaled α(l)", "#eeeeea", fs=8)
for y1, y2 in [(1.05, 1.3), (1.72, 1.95), (2.55, 2.8), (3.22, 3.45), (3.87, 4.1)]:
    arrow(ax, x0 + 1.2, y1, x0 + 1.2, y2)
arrow(ax, x0 + 2.4, 4.35, 7.0, 4.35)
x1 = 7.0
box(ax, x1, 4.1, 3.4, 0.5, "Pre-Norm  →  Router: top-8 of 256 experts", "#fde3d6", fs=8)
for j in range(8):
    on = j in (0, 2, 3, 6)
    box(ax, x1 + j * 0.42, 2.95, 0.36, 0.7, f"E{j+1}" if j < 7 else "…", BLUE if on else "#cfe0f5", fs=7,
        color="white" if on else INK)
ax.text(x1 - 0.05, 2.62, "256 fine-grained experts (8 active per token)", ha="left", fontsize=7, color=INK2)
box(ax, x1, 1.75, 3.4, 0.55, "Aux-loss-free balancing: per-expert bias b_i\n(bias-update step size on cosine decay)", "#fff1d1", fs=7.5)
box(ax, x1, 1.0, 3.4, 0.5, "Post-Norm  →  weighted sum  →  + residual", "#fde3d6", fs=8)
arrow(ax, x1 + 1.7, 4.1, x1 + 1.7, 3.68); arrow(ax, x1 + 3.2, 2.95, x1 + 3.2, 2.32)
arrow(ax, x1 + 1.7, 1.75, x1 + 1.7, 1.52)
ax.text(x1 + 1.7, 0.62, "Active per token ≈ 23.4B of 501.3B (≈4.7%)", ha="center", fontsize=8.5, color=BLUE, fontweight="bold")
ax.text(0.1, 0.05, "Sources: Reflection HF model card; vLLM recipe (SWA-512, NoPE full-attention); GGUF notes (9 of 52 full-attention layers). "
        "Hidden size not yet confirmed.", fontsize=7, color=INK2)
fig.savefig(f"{OUT}/fig1_architecture.png"); plt.close(fig)

# 2. Training pipeline -----------------------------------------------------
fig, ax = plt.subplots(figsize=(11, 3.6)); ax.set_xlim(0, 11); ax.set_ylim(0, 3.6); ax.axis("off")
ax.text(0.1, 3.35, "Beam training pipeline (as disclosed)", fontsize=13, fontweight="bold", color=INK)
stages = [
    ("Data curation", "~95% of raw web\ntokens discarded*\n152K-vocab tokenizer", "#eeeeea"),
    ("Pretraining", "23.8T tokens\n6,144 GB300 GPUs\n< 4 weeks", "#cfe0f5"),
    ("Mid-training", "long-context extension\nto 1M tokens\n(API beta: 262K)", "#cfe0f5"),
    ("Post-training (SFT)", "instruction + tool-use\nformatting (details\nnot disclosed)", "#d4f0e5"),
    ("Large-scale RL", "10,500 GB300, ~4 wks\n>100M rollouts, ~1M envs\nMirror CLI harness", BLUE),
]
w, gap = 1.95, 0.22
for i, (h, b, c) in enumerate(stages):
    x = 0.15 + i * (w + gap)
    dark = c == BLUE
    box(ax, x, 1.75, w, 0.6, h, c, fs=10, bold=True, color="white" if dark else INK)
    ax.text(x + w / 2, 1.55, b, ha="center", va="top", fontsize=8, color=INK2)
    if i < len(stages) - 1:
        arrow(ax, x + w, 2.05, x + w + gap, 2.05)
ax.text(0.15, 0.08, "* single-source figure (MarkTechPost). GPU counts and token totals per Reflection's launch post and launch coverage.",
        fontsize=7, color=INK2)
fig.savefig(f"{OUT}/fig2_pipeline.png"); plt.close(fig)

# 3. Total vs active parameters -------------------------------------------
models = ["Beam (Reflection)", "GLM-5.2 (Z.ai)", "Mistral Large 4", "DeepSeek V4 Pro", "Qwen 3.8 Max", "Kimi K3 (Moonshot)"]
total = [501, 753, 1000, 1600, 2400, 2800]
active = [23, 40, 49, 49, 95, 50]
fig, axes = plt.subplots(1, 2, figsize=(11, 3.9), sharey=True)
for ax, vals, lab, unit in [(axes[0], total, "Total parameters", "B"), (axes[1], active, "Active parameters per token", "B")]:
    cols = [BLUE if m.startswith("Beam") else GRAY for m in models]
    ax.barh(models, vals, color=cols, height=0.6)
    for y, v in enumerate(vals):
        ax.text(v + max(vals) * 0.015, y, f"{v:,}{unit}", va="center", fontsize=8.5, color=INK2)
    ax.set_title(lab, loc="left", fontsize=11, color=INK)
    ax.xaxis.grid(True, color=GRID); ax.set_axisbelow(True); ax.set_xlim(0, max(vals) * 1.18)
axes[0].invert_yaxis()
fig.suptitle("Open-weight MoE frontier: Beam is the smallest in both total and active size", x=0.01, ha="left",
             fontsize=13, fontweight="bold", color=INK)
fig.text(0.01, -0.03, "Sources: vendor model cards / launch coverage (Oct 2026). Kimi K3 active size disputed (~50B vs 104B); "
         "Mistral Large 4 ≈1T/49B per launch coverage. Beam: 501.3B / 23.4B per model card.", fontsize=7.5, color=INK2)
fig.tight_layout(); fig.savefig(f"{OUT}/fig3_params.png"); plt.close(fig)

# 4. Benchmarks vs open peers (Reflection's own table) ---------------------
import numpy as np
benches = ["SWE-Bench\nPro v1", "SWE-Bench Pro\nv2-Hard", "Terminal-\nBench 2.1", "GPQA\nDiamond", "HLE\n(no tools)", "AIME\n2026"]
data = {
    "Beam":         [65.5, 77.2, 80.1, 90.5, 36.2, 97.8],
    "GLM 5.2":      [62.1, None, 81.0, 91.2, 40.5, 99.2],
    "Kimi K3":      [None, 88.2, 88.3, 93.5, 46.9, None],
    "Qwen 3.8 Max": [67.7, None, 86.6, 92.6, 43.6, None],
}
colors = [BLUE, ORANGE, AQUA, YELLOW]
fig, ax = plt.subplots(figsize=(11, 4.6))
x = np.arange(len(benches)); bw = 0.2
for k, (name, vals) in enumerate(data.items()):
    for i, v in enumerate(vals):
        if v is None:
            ax.text(x[i] + (k - 1.5) * bw, 2, "NR", ha="center", fontsize=6.5, color=INK2, rotation=90)
            continue
        ax.bar(x[i] + (k - 1.5) * bw, v, bw * 0.9, color=colors[k], label=name if i == 0 or (k == 1 and i == 0) else None)
        if name == "Beam":
            ax.text(x[i] + (k - 1.5) * bw, v + 1.2, f"{v}", ha="center", fontsize=8, color=INK, fontweight="bold")
handles = [plt.Rectangle((0, 0), 1, 1, color=c) for c in colors]
ax.legend(handles, data.keys(), ncol=4, frameon=False, loc="upper left", bbox_to_anchor=(0, 1.02))
ax.set_xticks(x); ax.set_xticklabels(benches); ax.set_ylim(0, 108); ax.set_ylabel("Score (%)")
ax.yaxis.grid(True, color=GRID); ax.set_axisbelow(True)
title(ax, "Beam vs. leading open-weight peers (Reflection-reported, unverified)",
      "NR = not reported by Reflection for that model. Beam's values labelled.")
fig.savefig(f"{OUT}/fig4_bench_open.png"); plt.close(fig)

# 5. Beam vs frontier (small multiples) -----------------------------------
gpqa = [("Gemini 3.8 Flash", 95.6), ("GPT-5.4 Pro", 94.4), ("Gemini 3.1 Pro", 94.3), ("Kimi K3", 93.5),
        ("Claude Fable 5.1", 92.6), ("GLM 5.2", 91.2), ("Beam", 90.5), ("Llama 4 Maverick", 69.8)]
hle = [("Claude Fable 5.1", 60.9), ("Kimi K3", 46.9), ("Gemini 3.1 Pro", 44.4), ("Qwen 3.8 Max", 43.6),
       ("GPT-5.5", 41.4), ("GLM 5.2", 40.5), ("Beam", 36.2)]
tb = [("Gemini 3.8 Flash", 90.8), ("Kimi K3", 88.3), ("Qwen 3.8 Max", 86.6), ("GLM 5.2", 81.0), ("Beam", 80.1)]
fig, axes = plt.subplots(1, 3, figsize=(12, 4.4))
for ax, rows, t, lo in [(axes[0], gpqa, "GPQA Diamond", 60), (axes[1], hle, "HLE (no tools)", 0), (axes[2], tb, "Terminal-Bench 2.1", 60)]:
    names = [r[0] for r in rows][::-1]; vals = [r[1] for r in rows][::-1]
    ax.barh(names, vals, color=[BLUE if n == "Beam" else GRAY for n in names], height=0.62)
    for y, v in enumerate(vals):
        ax.text(v + 0.6, y, f"{v}", va="center", fontsize=8, color=INK2)
    ax.set_xlim(lo, 100 if lo else 70); ax.set_title(t, loc="left", fontsize=11, color=INK)
    ax.xaxis.grid(True, color=GRID); ax.set_axisbelow(True); ax.tick_params(axis="y", labelsize=8.5)
fig.suptitle("Beam vs. closed and open frontier models (best available public numbers)", x=0.01, ha="left",
             fontsize=13, fontweight="bold", color=INK)
fig.text(0.01, -0.04, "Mixed sources & dates: vendor cards, system cards and aggregators (Feb–Oct 2026). Axes truncated for "
         "GPQA/Terminal-Bench. Treat cross-vendor gaps < 2 pts as noise.", fontsize=7.5, color=INK2)
fig.tight_layout(); fig.savefig(f"{OUT}/fig5_bench_frontier.png"); plt.close(fig)

# 6. Efficiency: active params vs GPQA -------------------------------------
pts = [("Beam", 23, 90.5), ("GLM 5.2", 40, 91.2), ("Kimi K3", 50, 93.5), ("Qwen 3.8 Max", 95, 92.6)]
fig, ax = plt.subplots(figsize=(8, 4.4))
for n, a, g in pts:
    ax.scatter(a, g, s=110, color=BLUE if n == "Beam" else GRAY, edgecolor=SURF, linewidth=2, zorder=3)
    ax.annotate(n, (a, g), xytext=(8, -3), textcoords="offset points", fontsize=9, color=INK)
ax.set_xlabel("Active parameters per token (B)"); ax.set_ylabel("GPQA Diamond (%)")
ax.set_xlim(0, 110); ax.set_ylim(88, 95); ax.grid(True, color=GRID); ax.set_axisbelow(True)
title(ax, "Reasoning per active parameter", "Beam reaches ~97% of Kimi K3's GPQA score with < half the active compute per token")
fig.savefig(f"{OUT}/fig6_efficiency.png"); plt.close(fig)

# 7. Proposed research roadmap (Gantt) -------------------------------------
phases = [
    ("0. Reproduce & baseline (Beam weights, evals)", 0, 2),
    ("1. Data engine (curation, synthetic, envs)", 1, 6),
    ("2. Architecture ablations @1–7B", 1, 4),
    ("3. Scaling pilot ~60B-A3B (1–2T tokens)", 4, 3),
    ("4. Full pretrain ~500B-A20B (20–25T tokens)", 7, 4),
    ("5. Mid-train: long context 256K→1M", 11, 1.5),
    ("6. SFT + large-scale agentic RL", 12, 4),
    ("7. Safety, red-team, release", 15, 2),
]
fig, ax = plt.subplots(figsize=(11, 3.8))
for i, (n, s, d) in enumerate(phases):
    ax.barh(i, d, left=s, color=BLUE if i in (4, 6) else "#8db7e8", height=0.55)
ax.set_yticks(range(len(phases))); ax.set_yticklabels([p[0] for p in phases], fontsize=8.5); ax.invert_yaxis()
ax.set_xlabel("Month from project start"); ax.set_xlim(0, 18); ax.xaxis.grid(True, color=GRID); ax.set_axisbelow(True)
title(ax, "Proposed program to reach Beam-equivalent KPIs (~17 months)", "Darker bars = the two compute-dominant phases")
fig.savefig(f"{OUT}/fig7_roadmap.png"); plt.close(fig)


# 8. Independent intelligence index ---------------------------------------
aa = [("Claude Opus 5.5 (Anthropic)", 58, "c"), ("Claude Sonnet 5.5 (Anthropic)", 56, "c"), ("Claude Fable 5.1 (Anthropic)", 53, "c"),
      ("GPT-6 Astra (OpenAI)", 53, "c"), ("Gemini 4 Argon (Google)", 52.6, "c"), ("Grok 4.7 (xAI)", 46, "c"),
      ("MiMo-V2.6-Pro (Xiaomi)", 46, "o"), ("DeepSeek V4.1 Flash (DeepSeek)", 45, "o"), ("GLM-5.3 Max (Z.ai)", 45, "o"),
      ("Kimi K3 (Moonshot)", 44, "o"), ("Beam 1.0 (Reflection)", 30, "b"), ("Nemotron 3.5 Super VL (NVIDIA)", 26, "o")]
fig, ax = plt.subplots(figsize=(10, 5.2))
names = [a[0] for a in aa][::-1]; vals = [a[1] for a in aa][::-1]; kinds = [a[2] for a in aa][::-1]
cmap = {"c": "#8a8984", "o": "#c9c8c2", "b": BLUE}
ax.barh(names, vals, color=[cmap[k] for k in kinds], height=0.62)
for y, v in enumerate(vals):
    ax.text(v + 0.6, y, f"{v:g}", va="center", fontsize=8.5, color=INK2)
ax.set_xlim(0, 65); ax.xaxis.grid(True, color=GRID); ax.set_axisbelow(True); ax.tick_params(axis="y", labelsize=8.5)
handles = [plt.Rectangle((0, 0), 1, 1, color=cmap[k]) for k in ("c", "o", "b")]
ax.legend(handles, ["Closed (API only)", "Open weights", "Beam"], frameon=False, loc="lower right")
title(ax, "Independent ranking: Artificial Analysis Intelligence Index (v4.3.x, Oct 2026)",
      "One scale, one harness. Beam trails the closed frontier by ~28 points and the best open models by ~16")
fig.savefig(f"{OUT}/fig8_aa_index.png"); plt.close(fig)

# 9. Price vs intelligence -------------------------------------------------
pts = [("Claude Opus 5.5", 20, 58), ("Claude Sonnet 5.5", 10, 56), ("GPT-6 Astra / Claude Fable 5.1", 50, 53),
       ("Grok 4.7", 6, 46), ("DeepSeek V4.1 Flash", 0.28, 45)]
fig, ax = plt.subplots(figsize=(9, 4.6))
for n, pr, s_ in pts:
    ax.scatter(pr, s_, s=110, color="#8a8984" if "DeepSeek" not in n else "#c9c8c2", edgecolor=SURF, linewidth=2, zorder=3)
    ax.annotate(n, (pr, s_), xytext=(8, -3), textcoords="offset points", fontsize=8.5, color=INK)
ax.axhspan(29, 31, color=BLUE, alpha=0.15); ax.text(0.22, 31.6, "Beam 1.0 = 30 (open weights; self-host or partner pricing)", fontsize=8.5, color=BLUE)
ax.set_xscale("log"); ax.set_xlim(0.15, 400); ax.set_ylim(25, 62)
ax.set_xlabel("List output price, US$ per 1M tokens (log scale)"); ax.set_ylabel("AA Intelligence Index")
ax.grid(True, color=GRID); ax.set_axisbelow(True)
title(ax, "Intelligence vs. price: the market a new model enters", "Open models compete on cost; closed models on capability")
fig.savefig(f"{OUT}/fig9_price_vs_iq.png"); plt.close(fig)

# 10. Enterprise API share ------------------------------------------------
share = [("Anthropic", 40), ("OpenAI", 27), ("Google & all other closed", 22), ("Open-weight models", 11)]
fig, ax = plt.subplots(figsize=(9, 3.0))
names = [a for a, _ in share][::-1]; vals = [b for _, b in share][::-1]
ax.barh(names, vals, color=[AQUA if n.startswith("Open-weight") else GRAY for n in names], height=0.6)
for y, v in enumerate(vals):
    ax.text(v + 0.8, y, f"{v}%", va="center", fontsize=9, color=INK2)
ax.set_xlim(0, 50); ax.xaxis.grid(True, color=GRID); ax.set_axisbelow(True)
title(ax, "Enterprise LLM API spend share (Menlo Ventures, late 2025)", "Open-weight share fell from 19% to 11% in a year; 'Google & others' is the remainder")
fig.savefig(f"{OUT}/fig10_market_share.png"); plt.close(fig)

# 11. Strategy options cost -----------------------------------------------
opts = ["A. Adapt Beam\n(fine-tune + RL)", "B. Continued pretrain\n+ full post-train", "C. From-scratch\nBeam-class"]
comp = {"Compute": [3, 18, 55], "People (2 yrs)": [6, 25, 80], "Data & environments": [1, 8, 25], "Infra, eval, safety": [1, 5, 15]}
fig, ax = plt.subplots(figsize=(9, 4.2))
bottom = np.zeros(3); cols = [BLUE, ORANGE, AQUA, YELLOW]
for (k, v), c in zip(comp.items(), cols):
    ax.bar(opts, v, bottom=bottom, color=c, width=0.55, label=k, edgecolor=SURF, linewidth=2)
    bottom += np.array(v)
for i, t in enumerate(bottom):
    ax.text(i, t + 3, f"≈ ${t:.0f}M", ha="center", fontsize=10, fontweight="bold", color=INK)
ax.set_ylabel("US$ millions (illustrative midpoints)"); ax.set_ylim(0, 225)
ax.yaxis.grid(True, color=GRID); ax.set_axisbelow(True); ax.legend(frameon=False, ncol=4, loc="upper left")
title(ax, "Three ways to field a Beam-class model: indicative 2-year cost", "Author estimates; see cost model assumptions in Section 12")
fig.savefig(f"{OUT}/fig11_options_cost.png"); plt.close(fig)
print("done")
