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
ax.text(0.1, 6.95, "Beam 501B-A23B: architecture overview (reconstructed from public disclosures)",
        fontsize=13, fontweight="bold", color=INK)
ax.text(0.1, 6.65, "Confirmed by Reflection: sparse MoE, interleaved local/global attention, fine-grained routed experts, "
        "SandwichNorm,\nelementwise attention gating, depth-based residual scaling, FP32 residual accumulation, aux-loss-free balancing.",
        fontsize=8.5, color=INK2, va="top")
# left: stack
box(ax, 0.3, 0.4, 2.6, 0.55, "Token embedding (text-only)", "#dbe9f9", fs=9)
box(ax, 0.3, 5.15, 2.6, 0.55, "Final norm + LM head", "#dbe9f9", fs=9)
ax.add_patch(FancyBboxPatch((0.3, 1.2), 2.6, 3.7, boxstyle="round,pad=0.02,rounding_size=0.1", fc="#f1f0ec", ec=GRAY, lw=1))
ax.text(1.6, 4.65, "52 decoder layers*", ha="center", fontsize=9.5, fontweight="bold", color=INK)
pattern = ["L", "L", "L", "G"] * 3
for i, p in enumerate(pattern):
    y = 1.35 + i * 0.27
    box(ax, 0.55, y, 2.1, 0.21, ("Global attn + MoE" if p == "G" else "Local (sliding-window) attn + MoE"),
        BLUE if p == "G" else "#cfe0f5", fs=6.5, color="white" if p == "G" else INK)
ax.text(1.6, 1.24, "pattern repeats ... (ratio not officially confirmed*)", ha="center", fontsize=6.5, color=INK2)
arrow(ax, 1.6, 0.95, 1.6, 1.2); arrow(ax, 1.6, 4.9, 1.6, 5.15)
# right: block detail
ax.add_patch(FancyBboxPatch((3.6, 0.35), 7.1, 5.4, boxstyle="round,pad=0.02,rounding_size=0.1", fc="white", ec=GRAY, lw=1))
ax.text(3.8, 5.5, "One Beam decoder layer", fontsize=10.5, fontweight="bold", color=INK)
arrow(ax, 2.95, 3.0, 3.6, 3.0, color=GRAY)
x0 = 4.0
box(ax, x0, 0.6, 2.4, 0.45, "residual stream in (FP32 accum.)", "#eeeeea", fs=8)
box(ax, x0, 1.3, 2.4, 0.42, "Pre-Norm (SandwichNorm)", "#fde3d6", fs=8)
box(ax, x0, 1.95, 2.4, 0.6, "Self-attention\nlocal SWA or global (GQA*)", "#cfe0f5", fs=8)
box(ax, x0, 2.8, 2.4, 0.42, "Elementwise output gate  σ(xW_g)", "#d4f0e5", fs=8)
box(ax, x0, 3.45, 2.4, 0.42, "Post-Norm (SandwichNorm)", "#fde3d6", fs=8)
box(ax, x0, 4.1, 2.4, 0.5, "+ residual, depth-scaled α(l)", "#eeeeea", fs=8)
for y1, y2 in [(1.05, 1.3), (1.72, 1.95), (2.55, 2.8), (3.22, 3.45), (3.87, 4.1)]:
    arrow(ax, x0 + 1.2, y1, x0 + 1.2, y2)
arrow(ax, x0 + 2.4, 4.35, 7.0, 4.35)
# MoE side
x1 = 7.0
box(ax, x1, 4.1, 3.4, 0.5, "Pre-Norm  →  Router (learned, top-k)", "#fde3d6", fs=8)
for j in range(8):
    box(ax, x1 + j * 0.42, 2.9, 0.36, 0.75, f"E{j+1}", BLUE if j in (1, 4) else "#cfe0f5", fs=7,
        color="white" if j in (1, 4) else INK)
ax.text(x1 + 1.7, 2.7, "fine-grained routed experts (count not disclosed*)", ha="center", fontsize=7, color=INK2)
box(ax, x1, 1.85, 3.4, 0.5, "Aux-loss-free balancing: per-expert bias b_i\n(bias-update step size on cosine decay)", "#fff1d1", fs=7.5)
box(ax, x1, 1.05, 3.4, 0.5, "Post-Norm  →  weighted sum  →  + residual", "#fde3d6", fs=8)
arrow(ax, x1 + 1.7, 4.1, x1 + 1.7, 3.68); arrow(ax, x1 + 1.7, 2.9, x1 + 1.7, 2.37)
arrow(ax, x1 + 1.7, 1.85, x1 + 1.7, 1.57)
ax.text(x1 + 1.7, 0.62, "Active per token ≈ 23B of 501B (≈4.6%)", ha="center", fontsize=8.5, color=BLUE, fontweight="bold")
ax.text(0.1, 0.05, "* 52 layers reported by secondary coverage; 3:1 global/local reading, GQA and expert counts are analyst estimates "
        "pending the technical report.", fontsize=7, color=INK2)
fig.savefig(f"{OUT}/fig1_architecture.png"); plt.close(fig)

# 2. Training pipeline -----------------------------------------------------
fig, ax = plt.subplots(figsize=(11, 3.6)); ax.set_xlim(0, 11); ax.set_ylim(0, 3.6); ax.axis("off")
ax.text(0.1, 3.35, "Beam training pipeline (as disclosed)", fontsize=13, fontweight="bold", color=INK)
stages = [
    ("Data curation", "~95% of raw web\ntokens discarded*\ncode / STEM / agentic mix", "#eeeeea"),
    ("Pretraining", "23.8T tokens\n6,144 GB300 GPUs\n< 4 weeks", "#cfe0f5"),
    ("Mid-training", "long-context extension\nto 1M tokens\n(API beta: 262K)", "#cfe0f5"),
    ("Post-training (SFT)", "instruction + tool-use\nformatting (details\nnot disclosed)", "#d4f0e5"),
    ("Large-scale RL", "10,500 GB300, ~4 wks\n>100M rollouts, ~1M envs\n~1.3B sandboxes, 256K ctx", BLUE),
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
         "Mistral Large 4 ≈1T/49B per launch coverage.", fontsize=7.5, color=INK2)
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
print("done")
