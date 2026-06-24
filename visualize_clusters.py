"""
MRO Cluster Visualisation
--------------------------
Reduces sentence embeddings to 2D with UMAP, then renders an interactive
Plotly scatter plot. Each point = one unique task description.
Hover shows: description, cluster label, cluster size, min/max man hours.
"""

import re
import warnings
import numpy as np
import pandas as pd
import plotly.graph_objects as go
import umap
from sentence_transformers import SentenceTransformer
from sklearn.cluster import AgglomerativeClustering
from sklearn.preprocessing import normalize

warnings.filterwarnings("ignore")

# ── Config (must match mro_task_clustering.py) ────────────────────────────────
INPUT_FILE       = "Combined Sheet New Planes.xlsx"
OUTPUT_HTML      = "MRO_Cluster_Map.html"
DISTANCE_THRESHOLD = 0.35
EMBEDDING_MODEL  = "all-MiniLM-L6-v2"
# ─────────────────────────────────────────────────────────────────────────────


def clean_description(text: str) -> str:
    if not isinstance(text, str):
        return ""
    text = text.upper().strip()
    text = re.sub(r"[^A-Z0-9\s/\-]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def load_labour_rows(path: str) -> pd.DataFrame:
    print("Loading data …")
    df = pd.read_excel(path, sheet_name="Combined")
    df.columns = [re.sub(r'\s+', ' ', c).strip() for c in df.columns]
    df = df[df["Material / Services"].astype(str).str.strip().str.lower() == "labor cost"].copy()
    df = df[df["Description"].notna()].copy()
    df["Man Hour / Qty"] = pd.to_numeric(df["Man Hour / Qty"], errors="coerce")
    df = df[df["Man Hour / Qty"] > 0].copy()
    df["description_clean"] = df["Description"].apply(clean_description)
    return df.reset_index(drop=True)


def build_unique_desc_table(df: pd.DataFrame) -> pd.DataFrame:
    """One row per unique cleaned description, with aggregated stats."""
    agg = (
        df.groupby("description_clean")
        .agg(
            original_description=("Description", "first"),
            task_count=("Man Hour / Qty", "count"),
            min_mh=("Man Hour / Qty", "min"),
            mean_mh=("Man Hour / Qty", "mean"),
            max_mh=("Man Hour / Qty", "max"),
            tails=("Tail", lambda x: ", ".join(sorted(x.unique()))),
        )
        .reset_index()
    )
    return agg


def embed(descriptions: list[str]) -> np.ndarray:
    print(f"Embedding {len(descriptions):,} unique descriptions …")
    model = SentenceTransformer(EMBEDDING_MODEL)
    emb = model.encode(
        descriptions, batch_size=256, show_progress_bar=True, convert_to_numpy=True
    )
    return normalize(emb, norm="l2")


def cluster(embeddings: np.ndarray) -> np.ndarray:
    print("Clustering …")
    labels = AgglomerativeClustering(
        n_clusters=None,
        metric="cosine",
        linkage="average",
        distance_threshold=DISTANCE_THRESHOLD,
    ).fit_predict(embeddings)
    print(f"  {len(set(labels)):,} clusters")
    return labels


def reduce_to_2d(embeddings: np.ndarray) -> np.ndarray:
    print("Reducing to 2D with UMAP …")
    reducer = umap.UMAP(
        n_neighbors=15,
        min_dist=0.08,      # tighter packing — shows cluster shape clearly
        metric="cosine",
        random_state=42,
        verbose=False,
    )
    return reducer.fit_transform(embeddings)


def build_hover_text(row: pd.Series) -> str:
    desc = row["original_description"]
    # wrap long descriptions at 80 chars
    words = str(desc).split()
    lines, current = [], []
    for w in words:
        current.append(w)
        if len(" ".join(current)) > 80:
            lines.append(" ".join(current[:-1]))
            current = [w]
    lines.append(" ".join(current))
    wrapped = "<br>".join(lines)

    cluster_label = str(row["cluster_label"])[:90]
    return (
        f"<b>Cluster {row['cluster_id']}</b><br>"
        f"<i>{cluster_label}…</i><br><br>"
        f"<b>Description:</b><br>{wrapped}<br><br>"
        f"<b>Count:</b> {row['task_count']}  |  "
        f"<b>Man Hrs:</b> min {row['min_mh']:.1f} / max {row['max_mh']:.1f}<br>"
        f"<b>Aircraft:</b> {row['tails']}"
    )


def plot(df_unique: pd.DataFrame, coords: np.ndarray, out_path: str) -> None:
    print("Building interactive plot …")

    df_unique = df_unique.copy()
    df_unique["x"] = coords[:, 0]
    df_unique["y"] = coords[:, 1]
    df_unique["hover"] = df_unique.apply(build_hover_text, axis=1)

    # ── Colour by cluster size (log scale) so large clusters pop ─────────────
    df_unique["log_count"] = np.log1p(df_unique["task_count"])

    # ── Build one trace per cluster (enables toggling in legend) ─────────────
    # To keep the legend manageable, only show top-20 clusters by size;
    # everything else goes into a single "Other" trace.
    top_clusters = (
        df_unique.groupby("cluster_id")["task_count"]
        .sum()
        .nlargest(20)
        .index.tolist()
    )

    fig = go.Figure()

    # "Other" clusters — plotted first so they sit behind
    other = df_unique[~df_unique["cluster_id"].isin(top_clusters)]
    if len(other):
        fig.add_trace(go.Scatter(
            x=other["x"], y=other["y"],
            mode="markers",
            name="Other clusters",
            text=other["hover"],
            hovertemplate="%{text}<extra></extra>",
            marker=dict(
                size=5,
                color=other["log_count"],
                colorscale="Blues",
                opacity=0.45,
                showscale=False,
            ),
        ))

    # Top-20 clusters — each gets its own colour and legend entry
    colors = [
        "#e6194b","#3cb44b","#ffe119","#4363d8","#f58231",
        "#911eb4","#42d4f4","#f032e6","#bfef45","#fabed4",
        "#469990","#dcbeff","#9A6324","#fffac8","#800000",
        "#aaffc3","#808000","#ffd8b1","#000075","#a9a9a9",
    ]
    for i, cid in enumerate(top_clusters):
        grp = df_unique[df_unique["cluster_id"] == cid]
        label = str(grp["cluster_label"].iloc[0])[:50]
        fig.add_trace(go.Scatter(
            x=grp["x"], y=grp["y"],
            mode="markers",
            name=f"C{cid}: {label}",
            text=grp["hover"],
            hovertemplate="%{text}<extra></extra>",
            marker=dict(
                size=7 + grp["log_count"] * 1.5,  # bigger = appears more often
                color=colors[i % len(colors)],
                opacity=0.85,
                line=dict(width=0.4, color="white"),
            ),
        ))

    fig.update_layout(
        title=dict(
            text="MRO Labour Task Clusters — 2D Semantic Map",
            font=dict(size=20),
        ),
        xaxis=dict(title="UMAP-1", showgrid=False, zeroline=False),
        yaxis=dict(title="UMAP-2", showgrid=False, zeroline=False),
        plot_bgcolor="#0f1117",
        paper_bgcolor="#0f1117",
        font=dict(color="white"),
        legend=dict(
            title="Top 20 clusters (by count)",
            bgcolor="rgba(30,30,40,0.85)",
            bordercolor="#444",
            borderwidth=1,
            font=dict(size=10),
        ),
        hoverlabel=dict(bgcolor="#1e1e2e", font_size=12, font_color="white"),
        width=1400,
        height=900,
        margin=dict(l=40, r=40, t=60, b=40),
    )

    # ── Annotation explaining how to use ─────────────────────────────────────
    fig.add_annotation(
        text=(
            "Each point = one unique task description &nbsp;|&nbsp; "
            "Nearby points = semantically similar tasks &nbsp;|&nbsp; "
            "Hover for details &nbsp;|&nbsp; "
            "Click legend to hide/show clusters"
        ),
        xref="paper", yref="paper", x=0.5, y=-0.04,
        showarrow=False, font=dict(size=11, color="#aaa"),
        align="center",
    )

    fig.write_html(out_path, include_plotlyjs="cdn")
    print(f"Saved: {out_path}")
    print("Open it in any browser — all interactivity is built in.")


def main():
    import os
    os.chdir("/Users/bhumanpandita/Documents/IndiGo/MRO")

    df = load_labour_rows(INPUT_FILE)
    df_unique = build_unique_desc_table(df)

    descriptions = df_unique["description_clean"].tolist()
    embeddings   = embed(descriptions)
    labels       = cluster(embeddings)

    df_unique["cluster_id"] = labels

    # Cluster label = shortest description in cluster
    label_map = (
        df_unique.groupby("cluster_id")
        .apply(
            lambda g: g.loc[g["description_clean"].str.len().idxmin(), "original_description"],
            include_groups=False,
        )
        .to_dict()
    )
    df_unique["cluster_label"] = df_unique["cluster_id"].map(label_map)

    coords = reduce_to_2d(embeddings)
    plot(df_unique, coords, OUTPUT_HTML)


if __name__ == "__main__":
    main()
