# make_region_pie_final_english_legend_lines.py
# ---------------------------------------------
# Region pie chart with clear English legend:
# "Others include:" followed by one region per line.

from pathlib import Path
import math
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib import gridspec

# ===== SETTINGS =====
EXCEL_FILE = "data/elevate.xlsx"
SHEET_NAME = "data"
REGION_COL_NAME = "Region"
TOP_K = 12
OUTPUT_DIR = Path("output")
OUTPUT_PNG = OUTPUT_DIR / "region_pie_english_legend_lines.png"
TITLE = f"Region Distribution (Top {TOP_K} + Other)"

# Always grouped into "Other" (case-insensitive)
FORCE_OTHER = {
    "western greece",
    "central greece",
    "ionian islands",
    "south aegean",
    "north aegean",
    "epirus",
    "ήπειρος",
}
# ====================

def _find_region_col(df: pd.DataFrame) -> str:
    lookup = {c.strip().lower(): c for c in df.columns}
    key = REGION_COL_NAME.lower()
    if key not in lookup:
        raise ValueError(f"Column '{REGION_COL_NAME}' not found. Columns: {list(df.columns)}")
    return lookup[key]

def load_dataframe(xlsx_path: str, sheet_name=None) -> pd.DataFrame:
    if sheet_name is not None:
        return pd.read_excel(xlsx_path, sheet_name=sheet_name)
    xls = pd.ExcelFile(xlsx_path)
    for sh in xls.sheet_names:
        tmp = pd.read_excel(xlsx_path, sheet_name=sh, nrows=5)
        cols_lower = [c.strip().lower() for c in tmp.columns]
        if REGION_COL_NAME.lower() in cols_lower:
            return pd.read_excel(xlsx_path, sheet_name=sh)
    return pd.read_excel(xlsx_path, sheet_name=xls.sheet_names[0])

def normalize(x) -> str:
    if pd.isna(x):
        return ""
    return str(x).strip().lower()

def apply_forced_other(series: pd.Series) -> pd.Series:
    out = []
    for v in series:
        nv = normalize(v)
        if (nv in FORCE_OTHER) or (nv == ""):
            out.append("Other")
        else:
            out.append(v)
    return pd.Series(out, index=series.index)

def count_regions(series: pd.Series, top_k: int):
    """
    Επιστρέφει:
      - counts (Top-K + Other)
      - raw_other_values: ΛΙΣΤΑ ΜΕ ΤΑ ΑΡΧΙΚΑ ΟΝΟΜΑΤΑ περιοχών που μπήκαν στο Other,
        δηλ. forced + low-frequency/new + (blank)
    """
    # 1) Χαρτογράφηση: force (και κενά) -> "Other"
    mapped = apply_forced_other(series)

    # 2) Συνολικά counts στο χαρτογραφημένο
    counts_all = mapped.value_counts(dropna=False)

    # 3) Top-K από τα ΜΗ-Other
    keep = [lbl for lbl in counts_all.index if str(lbl) != "Other"][:top_k]

    # 4) Ό,τι ΔΕΝ είναι στα keep θεωρείται Other (συμπεριλαμβανομένου του ίδιου του "Other")
    is_other_mask = mapped.eq("Other") | ~mapped.isin(keep)
    other_count = int(is_other_mask.sum())

    kept_counts = counts_all.loc[keep]
    if other_count > 0:
        kept_counts = pd.concat([kept_counts, pd.Series({"Other": other_count})])

    # 5) ΛΙΣΤΑ ΓΙΑ LEGEND: από το ΑΡΧΙΚΟ series (ΟΧΙ από το mapped)
    raw_vals = series[is_other_mask]
    pretty = []
    for v in raw_vals:
        if pd.isna(v) or str(v).strip() == "":
            lab = "(blank)"
        else:
            lab = str(v).strip()
        pretty.append(lab)

    # μοναδικοποίηση + ωραία σειρά (πρώτα το (blank), μετά αλφαβητικά)
    raw_other_values = sorted(set(pretty), key=lambda x: (x != "(blank)", x.lower()))

    return kept_counts, raw_other_values

def _fmt_autopct(threshold=3.0):
    def inner(pct):
        return f"{pct:.1f}%" if pct >= threshold else ""
    return inner

def plot_pie_clean(ax, counts, title, min_pct=3.0):
    wedges, texts, autotexts = ax.pie(
        counts.values,
        labels=counts.index,
        autopct=_fmt_autopct(min_pct),
        startangle=90,
        pctdistance=0.75,
        labeldistance=1.10,
        wedgeprops=dict(linewidth=1, edgecolor="white"),
    )
    ax.set_title(title, pad=16)
    ax.axis("equal")
    for t in texts:
        t.set_fontsize(9)
    for t in autotexts:
        t.set_fontsize(9)

def make_figure(counts_ordered: pd.Series, raw_other: list[str], title: str, out_png: Path) -> None:
    fig = plt.figure(figsize=(17, 9))
    gs = gridspec.GridSpec(
        nrows=1, ncols=2, figure=fig,
        width_ratios=[1.0, 1.35],
        wspace=0.04
    )
    ax_pie = fig.add_subplot(gs[0, 0])
    ax_leg = fig.add_subplot(gs[0, 1])
    ax_leg.axis("off")

    plot_pie_clean(ax_pie, counts_ordered, title, min_pct=3.0)

    # --- Legend: clean English version, one region per line ---
    if raw_other:
        block = "Others include:\n" + "\n".join(raw_other)
        ax_leg.text(
            0.0, 1.0, block,
            fontsize=10.5, va="top", ha="left", wrap=True,
            bbox=dict(facecolor="white", edgecolor="lightgray", boxstyle="round,pad=0.5")
        )

    fig.savefig(out_png, dpi=200, bbox_inches="tight")
    plt.close(fig)

def main():
    OUTPUT_DIR.mkdir(exist_ok=True)
    df = load_dataframe(EXCEL_FILE, sheet_name=SHEET_NAME)
    region_col = _find_region_col(df)
    counts_ordered, raw_other_values = count_regions(df[region_col], top_k=TOP_K)
    make_figure(counts_ordered, raw_other_values, TITLE, OUTPUT_PNG)
    counts_ordered.to_csv(OUTPUT_DIR / "region_counts_topk_grouped_other.csv", header=["count"])
    with open(OUTPUT_DIR / "region_other_list.txt", "w", encoding="utf-8") as f:
        f.write("Others include:\n" + "\n".join(raw_other_values))
    print("Saved:", OUTPUT_PNG)
    print("Others include:")
    for item in raw_other_values:
        print(" -", item)

if __name__ == "__main__":
    main()
