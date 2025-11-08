# make_sector_pie_top14_legend.py
# --------------------------------
# Builds a Sector pie chart with your 14 short labels + groups the rest into "Other".
# Shows a separate legend panel listing exactly which original sectors are inside "Other".
#
# Output:
#   output/sector_pie_top14_legend.png
#   output/sector_counts_top14_other.csv
#   output/sector_other_list.txt

from pathlib import Path
import math
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib import gridspec

# ========= SETTINGS =========
EXCEL_FILE = "data/elevate.xlsx"     # change if needed
SHEET_NAME = "data"                  # set to None to auto-detect
SECTOR_COL_NAME = "Sector"           # case-insensitive
OUTPUT_DIR = Path("output")
OUTPUT_PNG = OUTPUT_DIR / "sector_pie_top14_legend.png"
TITLE = "Sector Distribution (Top 14 + Other)"
# ============================

# Your 14 concise labels (exactly as requested)
TARGET_LABELS = [
    "Life Sciences",
    "Environment & Energy",
    "Travel / Leisure",
    "Big Data and Analytics",
    "AgriTech/FoodTech",
    "Marketing",
    "Enterprise Software",
    "E-commerce and Fashion",
    "Manufacturing",
    "Financial Services",
    "Education",
    "Entertainment",
    "Logistics",
    "Real Estate",
]

# Map raw sector strings (as appear in the file) to the short labels above
SECTOR_MAPPING = {
    "Life Sciences (MedTech, HealthTech, BioTech)": "Life Sciences",
    "Environment & Energy (GreenTech, CleanTech)": "Environment & Energy",
    "Travel / Hospitality / Leisure": "Travel / Leisure",
    "Data Analytics - Big Data": "Big Data and Analytics",
    "AgriTech / FoodTech": "AgriTech/FoodTech",
    "Advertising & Marketing (AdTech)": "Marketing",
    "Enterprise Software": "Enterprise Software",
    "RetailTech – E-Commerce - FashionTech": "E-commerce and Fashion",
    "Manufacturing": "Manufacturing",
    "FinTech – Financial Services (WealthTech)": "Financial Services",
    "EduTech - Education": "Education",
    "Entertainment/Media (Games, Sports, Social)": "Entertainment",
    "Logistics & Transportation": "Logistics",
    "Real Estate (PropTech, Construction)": "Real Estate",
}


def _find_sector_col(df: pd.DataFrame) -> str:
    """Return the actual column name for SECTOR_COL_NAME (case-insensitive)."""
    lookup = {c.strip().lower(): c for c in df.columns}
    key = SECTOR_COL_NAME.lower()
    if key not in lookup:
        raise ValueError(f"Column '{SECTOR_COL_NAME}' not found. Columns: {list(df.columns)}")
    return lookup[key]


def load_dataframe(xlsx_path: str, sheet_name=None) -> pd.DataFrame:
    """Load the DataFrame. If sheet_name is None, auto-detect a sheet that contains the Sector column."""
    if sheet_name is not None:
        return pd.read_excel(xlsx_path, sheet_name=sheet_name)

    xls = pd.ExcelFile(xlsx_path)
    for sh in xls.sheet_names:
        tmp = pd.read_excel(xlsx_path, sheet_name=sh, nrows=5)
        cols_lower = [c.strip().lower() for c in tmp.columns]
        if SECTOR_COL_NAME.lower() in cols_lower:
            return pd.read_excel(xlsx_path, sheet_name=sh)

    # Fallback: first sheet
    return pd.read_excel(xlsx_path, sheet_name=xls.sheet_names[0])


def map_sectors_to_short(df: pd.DataFrame, sector_col_actual: str) -> pd.Series:
    """Map raw sector values to your short labels; anything unmapped becomes 'Other'."""
    short = df[sector_col_actual].map(SECTOR_MAPPING).fillna("Other")
    return short


def make_figure(counts_ordered: pd.Series, raw_other: list[str], title: str, out_png: Path) -> None:
    """Plot the pie on the left and a clean legend panel on the right (no overlaps, no truncation)."""
    # figure: wider canvas to fit a full legend comfortably
    fig = plt.figure(figsize=(16, 9))
    gs = gridspec.GridSpec(
        nrows=1, ncols=2, figure=fig,
        width_ratios=[1.0, 1.4],  # give extra width to legend panel
        wspace=0.08
    )

    ax_pie = fig.add_subplot(gs[0, 0])
    ax_leg = fig.add_subplot(gs[0, 1])
    ax_leg.axis("off")

    # --- Pie chart ---
    wedges, texts, autotexts = ax_pie.pie(
        counts_ordered.values,
        labels=counts_ordered.index,
        autopct="%1.1f%%",
        startangle=90,
    )
    ax_pie.set_title(title, pad=18)
    ax_pie.axis("equal")  # keep it circular

    # --- Legend: list what's inside 'Other' ---
    if raw_other:
        pretty = [f"– {item}" for item in raw_other]

        # If many lines, split into two columns to avoid wrapping ugliness
        if len(pretty) <= 20:
            text_block = "Other includes:\n" + "\n".join(pretty)
            ax_leg.text(
                0.0, 1.0, text_block,
                fontsize=10, va="top", ha="left", wrap=True,
                bbox=dict(facecolor="white", edgecolor="lightgray", boxstyle="round,pad=0.4")
            )
        else:
            mid = math.ceil(len(pretty) / 2)
            col1 = "\n".join(pretty[:mid])
            col2 = "\n".join(pretty[mid:])
            ax_leg.text(0.0, 1.0, "Other includes:", fontsize=11, va="top", ha="left")
            ax_leg.text(
                0.0, 0.97, col1,
                fontsize=9.5, va="top", ha="left",
                bbox=dict(facecolor="white", edgecolor="lightgray", boxstyle="round,pad=0.3")
            )
            ax_leg.text(
                0.52, 0.97, col2,
                fontsize=9.5, va="top", ha="left",
                bbox=dict(facecolor="white", edgecolor="lightgray", boxstyle="round,pad=0.3")
            )

    fig.savefig(out_png, dpi=200, bbox_inches="tight")
    plt.close(fig)


def main():
    OUTPUT_DIR.mkdir(exist_ok=True)

    # 1) Load and locate the sector column
    df = load_dataframe(EXCEL_FILE, sheet_name=SHEET_NAME)
    sector_col = _find_sector_col(df)

    # 2) Map to your 14 labels + 'Other'
    sector_short = map_sectors_to_short(df, sector_col)

    # 3) Counts and ordering (Top14 first, then Other, then any leftovers just in case)
    counts = sector_short.value_counts(dropna=False)
    ordered_labels = [lbl for lbl in TARGET_LABELS if lbl in counts.index]
    if "Other" in counts.index:
        ordered_labels.append("Other")
    leftovers = [lbl for lbl in counts.index if lbl not in ordered_labels]
    ordered_labels += leftovers
    counts_ordered = counts.reindex(ordered_labels)

    # 4) Which original raw sectors went to 'Other'?
    raw_other_values = df.loc[sector_short == "Other", sector_col].dropna().astype(str).unique().tolist()
    raw_other_values = sorted(set(raw_other_values))

    # 5) Plot (pie + right legend panel)
    make_figure(counts_ordered, raw_other_values, TITLE, OUTPUT_PNG)

    # 6) Save counts and the "Other" list to files
    counts_ordered.to_csv(OUTPUT_DIR / "sector_counts_top14_other.csv", header=["count"])
    with open(OUTPUT_DIR / "sector_other_list.txt", "w", encoding="utf-8") as f:
        f.write("Other includes:\n")
        for item in raw_other_values:
            f.write(f" - {item}\n")

    # Console summary
    print("Saved:", OUTPUT_PNG)
    print("Counts:")
    print(counts_ordered)
    if raw_other_values:
        print("\nOther includes:")
        for item in raw_other_values:
            print(" -", item)


if __name__ == "__main__":
    main()
