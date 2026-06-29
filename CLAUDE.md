# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Purpose

IndiGo MRO lease-return cost analysis. The goal is to cluster semantically similar aircraft maintenance task descriptions (written differently across MROs) into fine-grained groups, then identify the minimum man hours per task type to benchmark MRO pricing.

## Data

- **Input**: `Combined Sheet for ITPL (1).xlsx` — single sheet `Combined`, ~43K rows
- **Key columns**: `Tail` (aircraft ID, not MRO), `Description` (free-text task), `Material / Services` (filter for `Labor Cost`), `Man Hour / Qty`, `Card / WO`, `Sheet ref`
- **No MRO name column** — the output includes `Tail` + `Card / WO` so the user can trace rows back to the MRO manually
- `Man Hour / Qty` can contain `"FIXED COST"` (non-numeric) — always coerce with `errors="coerce"` and drop zeros/NaN

## Scripts

### `mro_task_clustering.py` — batch pipeline
Filters Labour Cost rows → embeds → clusters → exports Excel.

```bash
python3 mro_task_clustering.py
```

Output: `MRO_Labour_Task_Clusters.xlsx` with four sheets:
- `Search Clusters` / `Search Rows` — front-of-file keyword search sheets (see below)
- `Cluster Summary` — one row per cluster, min/mean/max man hours, Tail + Card/WO of the minimum row
- `All Labour Rows` — original rows with `cluster_id` and `cluster_label` appended

### `excel_search.py` — order-independent keyword search built into the Excel
For the ground team who work straight off the Excel file (no laptop / localhost).
`add_search_sheet()` inserts a sheet with one yellow input cell: typing keywords in
**any order** (e.g. `cockpit glass broken`) shows the matching rows below — it finds
`GLASS BROKEN IN COCKPIT` regardless of word order. Results update live as you type.
Requires Excel 365 / 2021+ (uses `TEXTSPLIT`).

**Do NOT use FILTER / dynamic-array formulas here.** openpyxl cannot write the OOXML
metadata that marks a formula as a true dynamic array, so Excel adds implicit-
intersection `@` operators on open (`=@IF`, `@SEARCH`, `@range`) which collapse every
array to one cell → `#VALUE!` or a single result. FILTER/REDUCE/BYROW/MMULT all failed
this way. The working design avoids dynamic arrays entirely:

- Two **hidden helper columns** on each data sheet: `__match` (1/0 per row, written as a
  legacy **array/CSE** `SUMPRODUCT` formula so `TEXTSPLIT` evaluates in array context
  without `@`) and `__rank` (running count among matches).
- The search sheet has a fixed block of `RESULT_SLOTS` (1000) result rows, each a plain
  `INDEX/MATCH` that pulls the i-th matching row by `__rank`. Plain per-row formulas are
  never broken by `@`.

Match logic: `SUMPRODUCT(--ISNUMBER(SEARCH(TEXTSPLIT(keywords), Description))) =
COLUMNS(TEXTSPLIT(keywords))` — every typed keyword found, case-insensitive, any order.
`save_output()` calls this for both data sheets. `TEXTSPLIT` needs the `_xlfn.` prefix
in storage or Excel shows `#NAME?`. To add the search sheets to an already-generated
workbook without re-clustering, run `python3 add_search.py "File.xlsx"` (idempotent —
`strip_helper_columns()` removes old helpers before re-adding).

### `cluster_explorer.py` — interactive Dash app (preferred)
Full dashboard: UMAP scatter map + click-to-inspect cluster panel + search.

```bash
python3 cluster_explorer.py
# Open http://127.0.0.1:8050
```

First launch computes embeddings + UMAP and caches to `cluster_data.pkl`. Subsequent launches load from cache instantly. **Delete `cluster_data.pkl` to force recomputation** (needed if threshold or input data changes).

### `visualize_clusters.py` — standalone HTML map
Generates `MRO_Cluster_Map.html` (static, no server needed). Less interactive than the Dash app.

```bash
python3 visualize_clusters.py
```

## Clustering Architecture

All three scripts share the same pipeline:

1. **Filter** — keep `Material / Services == "Labor Cost"`, drop zero/NaN man hours
2. **Clean** — uppercase, strip to `[A-Z0-9\s/\-]`, collapse whitespace
3. **Deduplicate** — embed only unique cleaned descriptions (avoids redundant computation)
4. **Embed** — `sentence-transformers` model `all-MiniLM-L6-v2`, L2-normalised
5. **Cluster** — `AgglomerativeClustering(metric="cosine", linkage="average", distance_threshold=0.35)`
6. **Reduce** (explorer/visualizer only) — UMAP to 2D for scatter plot
7. **Label** — cluster label = shortest description in the cluster

## Tuning

`DISTANCE_THRESHOLD = 0.35` controls cluster granularity — defined at the top of each script:
- Lower (e.g. `0.25`) → more clusters, finer grain
- Higher (e.g. `0.50`) → fewer clusters, broader groupings

After changing the threshold, delete `cluster_data.pkl` before re-running `cluster_explorer.py`.

## Dependencies

```
pandas, numpy, openpyxl, scipy, scikit-learn
sentence-transformers, torch
umap-learn, plotly, dash, pyarrow
```

Install missing packages with `pip3 install <pkg> --break-system-packages` (required on this macOS system due to PEP 668).
