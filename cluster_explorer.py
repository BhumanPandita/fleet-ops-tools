"""
MRO Cluster Explorer — Interactive Dash App
---------------------------------------------
Run:  python3 cluster_explorer.py
Then open:  http://127.0.0.1:8050

First launch computes embeddings + UMAP and caches to cluster_data.parquet.
Subsequent launches load instantly from cache.

Layout
------
LEFT   — UMAP scatter: click any point to select its cluster
RIGHT  — Cluster info card + scrollable task table
BOTTOM — Full-text search across all descriptions
"""

import os
import re
import warnings
import numpy as np
import pandas as pd
import plotly.graph_objects as go
from dash import Dash, Input, Output, State, callback_context, dcc, html, dash_table
import dash

warnings.filterwarnings("ignore")
os.chdir("/Users/bhumanpandita/Documents/IndiGo/MRO")

# ── Config ────────────────────────────────────────────────────────────────────
INPUT_FILE         = "Combined Sheet New Planes.xlsx"
CACHE_FILE         = "cluster_data.pkl"
DISTANCE_THRESHOLD = 0.35
EMBEDDING_MODEL    = "all-MiniLM-L6-v2"
# ─────────────────────────────────────────────────────────────────────────────

PALETTE = [
    "#4e79a7","#f28e2b","#e15759","#76b7b2","#59a14f",
    "#edc948","#b07aa1","#ff9da7","#9c755f","#bab0ac",
    "#00b4d8","#90e0ef","#caf0f8","#f72585","#7209b7",
    "#3a0ca3","#4361ee","#4cc9f0","#06d6a0","#ffd166",
    "#ef476f","#118ab2","#073b4c","#ffb703","#fb8500",
    "#8ecae6","#219ebc","#023047","#ffb703","#fd9e02",
]


# ── Data pipeline ─────────────────────────────────────────────────────────────

def clean(text):
    if not isinstance(text, str):
        return ""
    text = text.upper().strip()
    text = re.sub(r"[^A-Z0-9\s/\-]", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def build_dataset():
    print("Loading Excel …")
    df = pd.read_excel(INPUT_FILE, sheet_name="Combined")
    df.columns = [re.sub(r'\s+', ' ', c).strip() for c in df.columns]
    df = df[df["Material / Services"].astype(str).str.strip().str.lower() == "labor cost"].copy()
    df = df[df["Description"].notna()].copy()
    df["Man Hour / Qty"] = pd.to_numeric(df["Man Hour / Qty"], errors="coerce")
    df = df[df["Man Hour / Qty"] > 0].copy()
    df["desc_clean"] = df["Description"].apply(clean)

    # ── Per unique description: aggregate stats ───────────────────────────────
    agg = (
        df.groupby("desc_clean")
        .agg(
            orig_desc   = ("Description", "first"),
            task_count  = ("Man Hour / Qty", "count"),
            min_mh      = ("Man Hour / Qty", "min"),
            mean_mh     = ("Man Hour / Qty", "mean"),
            max_mh      = ("Man Hour / Qty", "max"),
            tails       = ("Tail", lambda x: ", ".join(sorted(x.unique()))),
            mros        = ("MRO",  lambda x: ", ".join(sorted(x.dropna().unique()))),
            cards       = ("Card / WO", lambda x: ", ".join(str(v) for v in x.unique())),
            sheet_refs  = ("Sheet ref", lambda x: ", ".join(sorted(x.dropna().unique()))),
        )
        .reset_index()
    )

    # ── Embed ─────────────────────────────────────────────────────────────────
    from sentence_transformers import SentenceTransformer
    from sklearn.cluster import AgglomerativeClustering
    from sklearn.preprocessing import normalize
    import umap

    print(f"Embedding {len(agg):,} unique descriptions …")
    model = SentenceTransformer(EMBEDDING_MODEL)
    emb = model.encode(
        agg["desc_clean"].tolist(),
        batch_size=256, show_progress_bar=True, convert_to_numpy=True,
    )
    emb = normalize(emb, norm="l2")

    print("Clustering …")
    labels = AgglomerativeClustering(
        n_clusters=None, metric="cosine",
        linkage="average", distance_threshold=DISTANCE_THRESHOLD,
    ).fit_predict(emb)
    agg["cluster_id"] = labels
    print(f"  {labels.max() + 1:,} clusters")

    print("UMAP → 2D …")
    coords = umap.UMAP(
        n_neighbors=15, min_dist=0.08,
        metric="cosine", random_state=42, verbose=False,
    ).fit_transform(emb)
    agg["x"] = coords[:, 0]
    agg["y"] = coords[:, 1]

    # ── Cluster label = shortest description in cluster ───────────────────────
    label_map = (
        agg.groupby("cluster_id")
        .apply(lambda g: g.loc[g["desc_clean"].str.len().idxmin(), "orig_desc"], include_groups=False)
        .to_dict()
    )
    agg["cluster_label"] = agg["cluster_id"].map(label_map).astype(str)

    agg.to_pickle(CACHE_FILE)
    print(f"Cached → {CACHE_FILE}")
    return agg


def load_data():
    if os.path.exists(CACHE_FILE):
        print(f"Loading cache from {CACHE_FILE} …")
        return pd.read_pickle(CACHE_FILE)
    return build_dataset()


# ── Load once at startup ───────────────────────────────────────────────────────
DF = load_data()

# Assign a display colour index per cluster (cycle through palette)
cluster_ids = sorted(DF["cluster_id"].unique())
color_map = {cid: PALETTE[i % len(PALETTE)] for i, cid in enumerate(cluster_ids)}
DF["color"] = DF["cluster_id"].map(color_map)

# Cluster-level summary (for the info card)
CLUSTER_SUMMARY = (
    DF.groupby("cluster_id")
    .agg(
        cluster_label = ("cluster_label", "first"),
        n_descriptions= ("orig_desc", "count"),
        total_tasks   = ("task_count", "sum"),
        min_mh        = ("min_mh", "min"),
        max_mh        = ("max_mh", "max"),
        mean_mh       = ("mean_mh", "mean"),
        mros          = ("mros", lambda x: ", ".join(sorted({m for s in x for m in s.split(", ") if m}))),
    )
    .reset_index()
)

# Per-cluster, per-MRO minimum man hours (used in detail panel)
MRO_BREAKDOWN = (
    DF.assign(mro_list=DF["mros"].str.split(", "))
    .explode("mro_list")
    .rename(columns={"mro_list": "mro"})
    .groupby(["cluster_id", "mro"])["min_mh"]
    .min()
    .reset_index()
    .rename(columns={"min_mh": "Min MH"})
)


# ── Base scatter (all points grey, no selection) ──────────────────────────────
def make_base_scatter(highlight_cid=None, search_text=""):
    df = DF.copy()

    if search_text:
        mask = df["orig_desc"].str.upper().str.contains(search_text.upper(), na=False)
        df["_alpha"] = np.where(mask, 1.0, 0.08)
        df["_size"]  = np.where(mask, 9, 4)
    elif highlight_cid is not None:
        df["_alpha"] = np.where(df["cluster_id"] == highlight_cid, 1.0, 0.06)
        df["_size"]  = np.where(df["cluster_id"] == highlight_cid, 11, 4)
    else:
        df["_alpha"] = 0.65
        df["_size"]  = 6

    fig = go.Figure()

    # background (non-selected) points as one trace for performance
    bg = df[df["_alpha"] < 0.5]
    if len(bg):
        fig.add_trace(go.Scattergl(
            x=bg["x"], y=bg["y"],
            mode="markers",
            name="",
            hoverinfo="skip",
            marker=dict(size=3, color="#334155", opacity=0.35),
            showlegend=False,
        ))

    # foreground (selected / search-match) points — per cluster for colour
    fg = df[df["_alpha"] >= 0.5]
    for cid, grp in fg.groupby("cluster_id"):
        col = color_map[cid]
        hover = [
            f"<b>Cluster {cid}</b><br>"
            f"{str(row.cluster_label)[:70]}<br><br>"
            f"<b>Description:</b> {str(row.orig_desc)[:120]}<br>"
            f"<b>Man Hrs:</b> min {row.min_mh:.1f} / max {row.max_mh:.1f}<br>"
            f"<b>MROs:</b> {row.mros}<br>"
            f"<b>Aircraft:</b> {row.tails}"
            for row in grp.itertuples()
        ]
        fig.add_trace(go.Scattergl(
            x=grp["x"], y=grp["y"],
            mode="markers",
            name=f"C{cid}",
            text=hover,
            hovertemplate="%{text}<extra></extra>",
            customdata=grp["cluster_id"].values,
            marker=dict(
                size=grp["_size"].values,
                color=col,
                opacity=0.92,
                line=dict(width=0.5, color="white"),
            ),
            showlegend=False,
        ))

    fig.update_layout(
        plot_bgcolor="#0f172a",
        paper_bgcolor="#0f172a",
        font=dict(color="#e2e8f0"),
        margin=dict(l=10, r=10, t=10, b=10),
        xaxis=dict(showgrid=False, zeroline=False, showticklabels=False, title=""),
        yaxis=dict(showgrid=False, zeroline=False, showticklabels=False, title=""),
        hovermode="closest",
        dragmode="zoom",
        uirevision="constant",   # keeps zoom/pan state across callbacks
    )
    return fig


# ── Dash app ──────────────────────────────────────────────────────────────────
app = Dash(__name__, suppress_callback_exceptions=True)

CARD_STYLE = dict(
    background="#1e293b", borderRadius="10px",
    padding="16px", marginBottom="12px",
)
LABEL_STYLE = dict(fontSize="11px", color="#94a3b8", marginBottom="4px")
VALUE_STYLE = dict(fontSize="15px", fontWeight="600", color="#f1f5f9")

app.layout = html.Div(
    style=dict(background="#0f172a", minHeight="100vh", fontFamily="'Inter',sans-serif", padding="16px"),
    children=[
        # ── Title bar ────────────────────────────────────────────────────────
        html.Div([
            html.H2("MRO Labour Task Cluster Explorer",
                    style=dict(color="#f1f5f9", margin="0 0 4px 0", fontSize="22px")),
            html.P(
                f"{len(DF):,} unique descriptions · {len(cluster_ids):,} clusters · click any point to explore",
                style=dict(color="#64748b", margin=0, fontSize="13px"),
            ),
        ], style=dict(marginBottom="16px")),

        # ── Search bar ───────────────────────────────────────────────────────
        html.Div([
            dcc.Input(
                id="search-box",
                type="text",
                placeholder="Search descriptions (e.g. 'cargo panel', 'engine mount') …",
                debounce=True,
                style=dict(
                    width="100%", padding="10px 14px", fontSize="14px",
                    background="#1e293b", color="#f1f5f9", border="1px solid #334155",
                    borderRadius="8px", outline="none", boxSizing="border-box",
                ),
            )
        ], style=dict(marginBottom="14px")),

        # ── Main row: scatter + detail panel ─────────────────────────────────
        html.Div(style=dict(display="flex", gap="16px"), children=[

            # Scatter plot
            html.Div(style=dict(flex="1 1 0"), children=[
                dcc.Graph(
                    id="scatter",
                    figure=make_base_scatter(),
                    config=dict(displayModeBar=True, scrollZoom=True),
                    style=dict(height="680px"),
                    clear_on_unhover=True,
                ),
            ]),

            # Detail panel
            html.Div(id="detail-panel", style=dict(width="440px", flexShrink="0"), children=[
                html.Div(
                    "Click any point on the map to explore its cluster.",
                    style=dict(color="#64748b", paddingTop="40px", textAlign="center", fontSize="14px"),
                )
            ]),
        ]),

        # Hidden store for selected cluster id
        dcc.Store(id="selected-cluster", data=None),
    ],
)


# ── Callbacks ─────────────────────────────────────────────────────────────────

@app.callback(
    Output("selected-cluster", "data"),
    Input("scatter", "clickData"),
    prevent_initial_call=True,
)
def store_click(click_data):
    if not click_data:
        return None
    pt = click_data["points"][0]
    cid = pt.get("customdata")
    return int(cid) if cid is not None else None


@app.callback(
    Output("scatter", "figure"),
    Input("selected-cluster", "data"),
    Input("search-box", "value"),
)
def update_scatter(selected_cid, search_text):
    return make_base_scatter(
        highlight_cid=selected_cid,
        search_text=(search_text or "").strip(),
    )


@app.callback(
    Output("detail-panel", "children"),
    Input("selected-cluster", "data"),
    Input("search-box", "value"),
)
def update_panel(selected_cid, search_text):
    search_text = (search_text or "").strip()

    # ── Search results mode ───────────────────────────────────────────────────
    if search_text and not selected_cid:
        matches = DF[DF["orig_desc"].str.upper().str.contains(search_text.upper(), na=False)]
        if matches.empty:
            return html.Div("No descriptions match your search.",
                            style=dict(color="#64748b", textAlign="center", paddingTop="40px"))
        rows = matches[["cluster_id","cluster_label","orig_desc","min_mh","max_mh","tails"]].copy()
        rows.columns = ["Cluster","Label","Description","Min MH","Max MH","Aircraft"]
        rows["Label"]       = rows["Label"].str[:60]
        rows["Description"] = rows["Description"].str[:100]
        return html.Div([
            html.Div(f"{len(matches):,} matching descriptions across {matches['cluster_id'].nunique()} clusters",
                     style=dict(color="#94a3b8", fontSize="13px", marginBottom="10px")),
            dash_table.DataTable(
                data=rows.to_dict("records"),
                columns=[{"name": c, "id": c} for c in rows.columns],
                style_table=dict(overflowY="auto", maxHeight="600px"),
                style_cell=dict(
                    background="#1e293b", color="#e2e8f0", border="1px solid #334155",
                    padding="8px", fontSize="12px", textAlign="left",
                    overflow="hidden", textOverflow="ellipsis", maxWidth="180px",
                ),
                style_header=dict(
                    background="#0f172a", color="#94a3b8",
                    fontWeight="600", border="1px solid #334155",
                ),
                tooltip_data=[
                    {c: {"value": str(row[c]), "type": "markdown"} for c in rows.columns}
                    for row in rows.to_dict("records")
                ],
                tooltip_delay=0, tooltip_duration=None,
                page_size=50,
            ),
        ])

    # ── Cluster detail mode ───────────────────────────────────────────────────
    if selected_cid is None:
        return html.Div(
            "Click any point on the map to explore its cluster.",
            style=dict(color="#64748b", paddingTop="40px", textAlign="center", fontSize="14px"),
        )

    summary = CLUSTER_SUMMARY[CLUSTER_SUMMARY["cluster_id"] == selected_cid]
    if summary.empty:
        return html.Div("Cluster not found.", style=dict(color="#ef4444"))

    s = summary.iloc[0]
    members = DF[DF["cluster_id"] == selected_cid].copy()

    # Build table of all descriptions in this cluster
    table_df = members[["orig_desc","min_mh","max_mh","mros","tails","cards"]].copy()
    table_df = table_df.sort_values("min_mh").reset_index(drop=True)
    table_df.columns = ["Description","Min MH","Max MH","MROs","Aircraft","Work Orders"]

    col = color_map.get(selected_cid, "#4e79a7")

    # Per-MRO min MH for this cluster
    mro_df = MRO_BREAKDOWN[MRO_BREAKDOWN["cluster_id"] == selected_cid].sort_values("Min MH")
    best_mh = mro_df["Min MH"].min() if not mro_df.empty else None

    return html.Div([
        # Cluster header
        html.Div([
            html.Div(f"Cluster {selected_cid}", style=dict(fontSize="12px", color=col, fontWeight="700", marginBottom="4px")),
            html.Div(str(s["cluster_label"])[:120], style=dict(fontSize="14px", color="#f1f5f9", fontWeight="600", lineHeight="1.4")),
        ], style=CARD_STYLE),

        # Stats row
        html.Div(style=dict(**CARD_STYLE, display="grid", gridTemplateColumns="1fr 1fr 1fr", gap="12px"), children=[
            html.Div([html.Div("Unique Descriptions", style=LABEL_STYLE), html.Div(str(s["n_descriptions"]), style=VALUE_STYLE)]),
            html.Div([html.Div("Total Task Rows", style=LABEL_STYLE), html.Div(str(int(s["total_tasks"])), style=VALUE_STYLE)]),
            html.Div([html.Div("Min Man Hours", style=LABEL_STYLE), html.Div(f"{s['min_mh']:.1f}", style=dict(**VALUE_STYLE, color="#4ade80"))]),
            html.Div([html.Div("Mean Man Hours", style=LABEL_STYLE), html.Div(f"{s['mean_mh']:.1f}", style=VALUE_STYLE)]),
            html.Div([html.Div("Max Man Hours", style=LABEL_STYLE), html.Div(f"{s['max_mh']:.1f}", style=dict(**VALUE_STYLE, color="#f87171"))]),
        ]),

        # Per-MRO min man hours — the key comparison
        html.Div("Min Man Hours by MRO", style=dict(color="#94a3b8", fontSize="12px", marginBottom="8px", fontWeight="600")),
        html.Div(
            style=dict(display="flex", gap="10px", flexWrap="wrap", marginBottom="12px"),
            children=[
                html.Div([
                    html.Div(row["mro"], style=dict(fontSize="11px", color="#94a3b8", marginBottom="3px")),
                    html.Div(f"{row['Min MH']:.1f}", style=dict(
                        fontSize="22px", fontWeight="700",
                        color="#4ade80" if row["Min MH"] == best_mh else "#f1f5f9",
                    )),
                    html.Div("MH", style=dict(fontSize="10px", color="#64748b")),
                ], style=dict(
                    background="#0f172a", borderRadius="8px", padding="10px 18px",
                    border=f"2px solid {'#4ade80' if row['Min MH'] == best_mh else '#334155'}",
                    textAlign="center",
                ))
                for _, row in mro_df.iterrows()
            ]
        ) if not mro_df.empty else html.Div(),

        # All descriptions table
        html.Div("All Descriptions in this Cluster", style=dict(color="#94a3b8", fontSize="12px", marginBottom="8px", fontWeight="600")),
        dash_table.DataTable(
            data=table_df.to_dict("records"),
            columns=[{"name": c, "id": c} for c in table_df.columns],
            style_table=dict(overflowY="auto", maxHeight="380px"),
            style_cell=dict(
                background="#1e293b", color="#e2e8f0", border="1px solid #334155",
                padding="8px 10px", fontSize="12px", textAlign="left",
                overflow="hidden", textOverflow="ellipsis", maxWidth="200px",
            ),
            style_header=dict(
                background="#0f172a", color="#94a3b8",
                fontWeight="600", border="1px solid #334155", fontSize="11px",
            ),
            style_data_conditional=[
                {"if": {"row_index": 0}, "background": "#14532d", "borderLeft": f"3px solid #4ade80"},
            ],
            tooltip_data=[
                {"Description": {"value": row["Description"], "type": "markdown"}}
                for row in table_df.to_dict("records")
            ],
            tooltip_delay=0, tooltip_duration=None,
            page_size=30,
        ),
        html.Div("First row (green) = minimum man hours",
                 style=dict(color="#4ade80", fontSize="11px", marginTop="6px")),
    ])


if __name__ == "__main__":
    print("\n  Open your browser at:  http://127.0.0.1:8050\n")
    app.run(debug=False, port=8050)
