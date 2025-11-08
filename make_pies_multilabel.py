import re
from pathlib import Path
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import numpy as np

# ================= USER SETTINGS =================
EXCEL_FILE = "elevate.xlsx"   # change if needed
SHEET_NAME = "data"           # set to None to auto-detect
TECH_COL = "Technology"       # column to parse

OUTPUT_DIR = Path("output")
MAKE_ONE_HOT_COLUMNS = True   # also save a sheet with 0/1 columns per category
# =================================================

# Canonical buckets you asked for (exact names used in the final pie)
BUCKETS = [
    "AI",
    "Data Analytics",
    "Big Data",
    "IoT",
    "Mobile Applications",
    "Software",
    "Cloud computing",
    "Other",
]

# Case-insensitive keyword patterns for each bucket (tweak as you like)
# We intentionally keep these simple and robust to punctuation & lists.
PATTERNS = {
    "AI": [
        r"\bAI\b",
        r"artificial intelligence",
        r"\bmachine learning\b",   # include ML if you want it counted under AI
    ],
    "Data Analytics": [
        r"data analytics",
        r"\banalytics\b",
        r"data analysis",
    ],
    "Big Data": [
        r"\bbig data\b",
    ],
    "IoT": [
        r"\bIoT\b",
        r"internet of things",
    ],
    "Mobile Applications": [
        r"web or mobile application",
        r"\bmobile application",
        r"\bmobile app(s)?\b",
        r"\bweb application\b",
        r"\bapp(s)?\b",
    ],
    "Software": [
        r"\bsoftware\b",
    ],
    "Cloud computing": [
        r"cloud computing",
        r"\bcloud\b",
    ],
    # "Other" handled specially below
}

def normalize_text(s):
    if pd.isna(s):
        return ""
    return str(s).strip().lower()

def row_to_buckets(text):
    """
    Return set of buckets matched in this text.
    - Matches ALL that apply (multi-label).
    - Adds 'Other' if 'other' literal appears OR if nothing else matched.
    """
    t = normalize_text(text)
    matched = set()
    if not t:
        # empty => counts as Other
        return {"Other"}

    # match each defined bucket
    for bucket, patterns in PATTERNS.items():
        for pat in patterns:
            if re.search(pat, t, flags=re.IGNORECASE):
                matched.add(bucket)
                break

    # explicit 'Other' in text -> also count Other
    if "other" in t:
        matched.add("Other")

    # if nothing matched at all -> count as Other
    if not matched:
        matched.add("Other")

    return matched

def main():
    OUTPUT_DIR.mkdir(exist_ok=True)

    # Load sheet
    if SHEET_NAME is None:
        # Auto-detect a sheet that has the tech column
        xls = pd.ExcelFile(EXCEL_FILE)
        chosen = None
        for sh in xls.sheet_names:
            tmp = pd.read_excel(EXCEL_FILE, sheet_name=sh)
            if any(c.strip().lower() == TECH_COL.lower() for c in tmp.columns):
                chosen = sh
                df = tmp
                break
        if chosen is None:
            # Fallback to first sheet
            df = pd.read_excel(EXCEL_FILE, sheet_name=xls.sheet_names[0])
            chosen = xls.sheet_names[0]
    else:
        df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)
        chosen = SHEET_NAME

    # Find actual Technology column name (case-insensitive)
    col_lookup = {c.strip().lower(): c for c in df.columns}
    if TECH_COL.lower() not in col_lookup:
        raise ValueError(f"Column '{TECH_COL}' not found in sheet '{chosen}'. Columns: {list(df.columns)}")
    tech_col = col_lookup[TECH_COL.lower()]

    # Multi-label counting
    counts = {b: 0 for b in BUCKETS}
    one_hot_rows = []  # for optional one-hot output

    for _, row in df.iterrows():
        buckets = row_to_buckets(row[tech_col])

        # increment each matched bucket
        for b in buckets:
            if b in counts:
                counts[b] += 1
            else:
                # shouldn't happen, but keep safe
                counts["Other"] += 1

        if MAKE_ONE_HOT_COLUMNS:
            one_hot = {b: (1 if b in buckets else 0) for b in BUCKETS}
            one_hot_rows.append(one_hot)

    # Create a DataFrame of counts in a fixed order
    counts_series = pd.Series([counts[b] for b in BUCKETS], index=BUCKETS, name="count")
    counts_df = counts_series.to_frame()

    # Save counts to Excel
    with pd.ExcelWriter(OUTPUT_DIR / "tech_multilabel_counts.xlsx", engine="openpyxl") as writer:
        counts_df.to_excel(writer, sheet_name="counts")
        if MAKE_ONE_HOT_COLUMNS:
            oh = pd.DataFrame(one_hot_rows, columns=BUCKETS) if one_hot_rows else pd.DataFrame(columns=BUCKETS)
            # Include the original Company and Technology columns for reference if present
            keep_cols = []
            for k in ["Company", tech_col]:
                if k in df.columns and k not in keep_cols:
                    keep_cols.append(k)
            out_df = pd.concat([df[keep_cols].reset_index(drop=True), oh.reset_index(drop=True)], axis=1)
            out_df.to_excel(writer, sheet_name="one_hot", index=False)

    # Plot pie chart for these 8 buckets
    total = counts_series.sum()
    if total == 0:
        print("No data to plot.")
        return

    # Sort by size descending (optional: comment this line to keep fixed order)
    counts_series_sorted = counts_series.sort_values(ascending=False)

    # Save PNG and combined PDF
    png_path = OUTPUT_DIR / "pie_technology_multilabel.png"
    pdf_path = OUTPUT_DIR / "pie_technology_multilabel.pdf"

    # Single pie with all 8 buckets
    plt.figure(figsize=(8, 8))
    plt.pie(
        counts_series_sorted.values,
        labels=counts_series_sorted.index,
        autopct="%1.1f%%",
        startangle=90,
    )
    plt.title(f"Technology categories (multi-label count) — sheet: {chosen}")
    try:
        plt.tight_layout()
    except Exception:
        pass
    plt.savefig(png_path, dpi=160)
    with PdfPages(pdf_path) as pdf:
        pdf.savefig()
    plt.close()

    print("Done!")
    print("Counts:")
    print(counts_df)
    print("Saved files:")
    print(" -", OUTPUT_DIR / "tech_multilabel_counts.xlsx")
    print(" -", png_path)
    print(" -", pdf_path)

if __name__ == "__main__":
    main()
