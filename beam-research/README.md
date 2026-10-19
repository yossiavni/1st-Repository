# Building a Beam-Class LLM: research baseline and executive deck

This folder holds research for building a new LLM comparable to Reflection AI's Beam 1.0, with a business case for competing against Anthropic, OpenAI, Google, xAI, Meta and the open-weight leaders. Sources go up to 19 Oct 2026.

| File | What it is |
|---|---|
| `Beam_Research_Baseline.docx` | Full research baseline, v2.0 (37 pages, 11 figures, 42 sources). In Word, right-click the table of contents and choose *Update Field*. |
| `Beam_Research_Baseline.pdf` | PDF rendering of the baseline. |
| `Beam_Executive_Summary.pptx` | 16-slide executive deck with speaker notes. |
| `figures/` | Architecture and pipeline diagrams, benchmark, market and cost charts. |
| `src/make_figures.py` | Regenerates the figures: `python3 src/make_figures.py figures` (needs matplotlib). |
| `src/build_docx.js`, `src/docx_helpers.js` | Rebuild the Word document: `node src/build_docx.js` (needs the `docx` npm package). |
| `src/build_pptx.js` | Rebuilds the deck: `NODE_PATH=<node_modules with pptxgenjs, react-icons, react, react-dom, sharp> node src/build_pptx.js <pptx-skill-dir>`. The second argument is the directory that holds `scripts/apply_theme.js`. |

Confidence tags used throughout: [V] vendor-reported, [I] independently measured, [E] author estimate, [S] single secondary source.
