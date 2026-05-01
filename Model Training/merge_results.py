#!/usr/bin/env python3
"""
merge_results_comprehensive_final.py
=============================================================================
THE COMPLETE UNABRIDGED THESIS ANALYTICS ENGINE
=============================================================================
Filipino Fake News Detection — Full Statistical & Visual Pipeline.

OUTPUT FILES
------------
Thesis_Results_Final_Comprehensive.xlsx
    {CondKey}_Overall    — Binary metrics per model
    {CondKey}_Subclass   — Per-subclass accuracy
    {CondKey}_Stats      — Within-condition McNemar (same arch, vary HF:AIF)
    CrossCondition_Stats — Cross-condition McNemar (same arch+HF, vary HR:AIR)

Thesis_Champion_Comparison.xlsx
    Champion_Comparisons — Post-hoc top-3 per (condition × architecture)

MCNEMAR SORT ORDER (all three tables)
--------------------------------------
Primary   : Architecture  — Tagalog BERT (0) before Tagalog DistilBERT (1)
Secondary : HF% of left model  — descending numeric  (100 → 75 → 50 → 25 → 0)
Tertiary  : HF% of right model — descending numeric
Quaternary: Condition pair or accuracy pair (where applicable)
"""

import argparse
import os
import re
import warnings
from typing import Tuple

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import seaborn as sns
from sklearn.metrics import precision_recall_fscore_support, accuracy_score

plt.rcParams['figure.dpi']      = 300
plt.rcParams['savefig.dpi']     = 300
plt.rcParams['font.family']     = 'sans-serif'
plt.rcParams['font.sans-serif'] = ['Arial']

warnings.filterwarnings("ignore")

# =============================================================================
# 1. GLOBAL CONSTANTS
# =============================================================================

SUBCLASSES = ["HR", "AI-R", "HF", "AI-F"]

CONDITIONS = {
    "Condition_A": "HR100",
    "Condition_B": "HR67",
    "Condition_C": "HR50",
}
CONDITION_LABELS = {
    "Condition_A": "Condition A: Human-Real 100%",
    "Condition_B": "Condition B: Human-Real 67%",
    "Condition_C": "Condition C: Human-Real 50%",
}
CONDITION_SHORT = {
    "Condition_A": "CondA",
    "Condition_B": "CondB",
    "Condition_C": "CondC",
}
# Canonical display order for conditions (used as sort index)
CONDITION_ORDER = {v: i for i, v in enumerate(CONDITION_SHORT.values())}
# Canonical display order for architectures (Tagalog BERT first)
ARCH_ORDER = {"Tagalog BERT": 0, "Tagalog DistilBERT": 1}

LABEL_STR_TO_INT = {"real": 0, "fake": 1}

# =============================================================================
# 2. LABEL HELPERS
# =============================================================================

def extract_architecture(model_key: str) -> str:
    """Returns 'Tagalog DistilBERT' or 'Tagalog BERT'."""
    return "Tagalog DistilBERT" if "distilbert" in model_key.lower() else "Tagalog BERT"


def extract_hf_label(model_key: str) -> str:
    """Returns 'HF{n}-AIF{m}', inferring AIF = 100 - HF when absent."""
    key   = model_key.lower()
    hf_m  = re.search(r'hf(\d+)',  key)
    aif_m = re.search(r'aif(\d+)', key)
    hf    = int(hf_m.group(1))  if hf_m  else 0
    aif   = int(aif_m.group(1)) if aif_m else (100 - hf)
    return f"HF{hf}-AIF{aif}"


def extract_hf_pct(model_key_or_hf_label: str) -> int:
    """
    Extracts the numeric HF percentage from either a model key or an
    HF label string such as 'HF75-AIF25'.  Returns 0 if not found.
    Used exclusively for *sort keys* — never shown to the user.
    """
    m = re.search(r'hf(\d+)', model_key_or_hf_label, re.IGNORECASE)
    return int(m.group(1)) if m else 0


def format_model_display_label(model_key: str) -> str:
    """
    Full publication label.
    e.g. "bert-HR67-HF100" -> "Tagalog BERT: HR67-AIR33-HF100-AIF0"
    """
    key  = model_key.lower()
    arch = "Tagalog DistilBERT" if "distilbert" in key else "Tagalog BERT"

    def _get(pattern):
        m = re.search(pattern, key, re.IGNORECASE)
        return int(m.group(1)) if m else None

    hr  = _get(r'(?<!ai)hr(\d+)')
    air = _get(r'air(\d+)')
    hf  = _get(r'hf(\d+)')
    aif = _get(r'aif(\d+)')

    if hr  is not None and air is None: air = 100 - hr
    elif air is not None and hr is None: hr  = 100 - air
    if hf  is not None and aif is None: aif = 100 - hf
    elif aif is not None and hf is None: hf  = 100 - aif

    hr  = hr  or 0
    air = air or 0
    hf  = hf  or 0
    aif = aif or 0

    return f"HR{hr}-AIR{air}-HF{hf}-AIF{aif}"


def compute_logical_sort_key(model_identifier: str) -> Tuple:
    """Tagalog BERT before Tagalog DistilBERT; HF ratio descending within each arch."""
    mid   = model_identifier.lower()
    arch  = 1 if "distilbert" in mid else 0
    hf_m  = re.search(r'hf(\d+)', mid)
    hf    = int(hf_m.group(1)) if hf_m else 0
    return (arch, -hf)


def standardize_and_sort_thesis_data(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    if 'Model Identifier' in df.columns:
        # 1. Compute sort key FIRST using the raw key (to detect architecture)
        df['_sort_key'] = df['Model Identifier'].apply(compute_logical_sort_key)
        
        # 2. Format the display label SECOND (removing the prefix)
        df['Model Identifier'] = df['Model Identifier'].apply(format_model_display_label)
        
        # 3. Sort by the hidden key and then drop it
        df = df.sort_values('_sort_key').drop(columns=['_sort_key'])
    return df

# =============================================================================
# 3. SHARED MCNEMAR HELPERS
# =============================================================================

def _run_mcnemar_pair(res_1: pd.DataFrame, res_2: pd.DataFrame):
    """Single McNemar test between two result DataFrames. Returns p-value or None."""
    from statsmodels.stats.contingency_tables import mcnemar

    n = min(len(res_1), len(res_2))
    if n == 0:
        return None

    gt = res_1['true_int'].values[:n]
    p1 = res_1['pred_int'].values[:n]
    p2 = res_2['pred_int'].values[:n]

    b = int(np.sum((p1 == gt) & (p2 != gt)))
    c = int(np.sum((p1 != gt) & (p2 == gt)))

    return mcnemar([[0, b], [c, 0]], exact=(b + c < 25)).pvalue


def _apply_bh(df: pd.DataFrame) -> pd.DataFrame:
    """Adds BH-corrected p-values and significance flag. Does NOT sort."""
    from statsmodels.stats.multitest import multipletests

    if df.empty:
        return df
    _, p_adj, _, _ = multipletests(df['p-value'], alpha=0.05, method='fdr_bh')
    df = df.copy()
    df['p-adj (BH)']  = np.round(p_adj, 4)
    df['Significant'] = p_adj < 0.05
    return df


def _arch_sort_val(arch_str: str) -> int:
    """Tagalog BERT → 0, Tagalog DistilBERT → 1."""
    return 0 if "distilbert" not in arch_str.lower() else 1

# =============================================================================
# 4. WITHIN-CONDITION MCNEMAR
# =============================================================================

def run_within_condition_mcnemar(condition_df: pd.DataFrame) -> pd.DataFrame:
    """
    Same architecture, same HR:AIR (fixed by condition) -> all HF:AIF pairs.
    Cross-architecture pairs are excluded.

    Sort order
    ----------
    1. Architecture  : Tagalog BERT first, Tagalog DistilBERT second
    2. HF% of left model  : descending (100 → 75 → 50 → 25 → 0)
    3. HF% of right model : descending
    """
    models = condition_df['model_key'].unique()
    rows   = []

    for i in range(len(models)):
        for j in range(i + 1, len(models)):
            m1, m2 = models[i], models[j]
            if extract_architecture(m1) != extract_architecture(m2):
                continue

            res_1 = condition_df[condition_df['model_key'] == m1].sort_values('sample_index')
            res_2 = condition_df[condition_df['model_key'] == m2].sort_values('sample_index')

            pval = _run_mcnemar_pair(res_1, res_2)
            if pval is None:
                continue

            hf_label_1 = extract_hf_label(m1)
            hf_label_2 = extract_hf_label(m2)

            # Always put the higher-HF model on the left side of "vs"
            if extract_hf_pct(hf_label_1) < extract_hf_pct(hf_label_2):
                hf_label_1, hf_label_2 = hf_label_2, hf_label_1

            rows.append({
                "Architecture":  extract_architecture(m1),
                "Comparison":    f"{hf_label_1} vs {hf_label_2}",
                "p-value":       round(pval, 4),
                # Numeric sort keys (dropped before returning)
                "_arch_order":   _arch_sort_val(extract_architecture(m1)),
                "_hf_left":      -extract_hf_pct(hf_label_1),   # negative → descending
                "_hf_right":     -extract_hf_pct(hf_label_2),
            })

    out = _apply_bh(pd.DataFrame(rows))
    if out.empty:
        return out

    out = (out
           .sort_values(["_arch_order", "_hf_left", "_hf_right"])
           .drop(columns=["_arch_order", "_hf_left", "_hf_right"])
           .reset_index(drop=True))
    return out

# =============================================================================
# 5. CROSS-CONDITION MCNEMAR
# =============================================================================

def run_cross_condition_mcnemar(total_df: pd.DataFrame) -> pd.DataFrame:
    """
    Same architecture + HF:AIF -> all condition pairs (A<->B, A<->C, B<->C).

    Sort order
    ----------
    1. Architecture      : Tagalog BERT first, Tagalog DistilBERT second
    2. HF% (descending)  : HF100 → HF75 → HF50 → HF25 → HF0
    3. Condition pair    : CondA vs CondB → CondA vs CondC → CondB vs CondC
    """
    df = total_df.copy()

    def _tag(mk):
        ml = mk.lower()
        for ck, hr in CONDITIONS.items():
            if hr.lower() in ml:
                return ck
        return None

    df['_cond']     = df['model_key'].apply(_tag)
    df['_arch']     = df['model_key'].apply(extract_architecture)
    df['_hf_label'] = df['model_key'].apply(extract_hf_label)

    cond_keys = list(CONDITIONS.keys())
    rows      = []

    for (arch, hf_label), grp in df.groupby(['_arch', '_hf_label']):
        for ci in range(len(cond_keys)):
            for cj in range(ci + 1, len(cond_keys)):
                ck_i, ck_j = cond_keys[ci], cond_keys[cj]

                r_i = grp[grp['_cond'] == ck_i].sort_values('sample_index')
                r_j = grp[grp['_cond'] == ck_j].sort_values('sample_index')

                if r_i.empty or r_j.empty:
                    continue

                pval = _run_mcnemar_pair(r_i, r_j)
                if pval is None:
                    continue

                short_i = CONDITION_SHORT[ck_i]
                short_j = CONDITION_SHORT[ck_j]

                rows.append({
                    "Architecture":  arch,
                    "Comparison":    f"{hf_label} [{short_i}] vs [{short_j}]",
                    "p-value":       round(pval, 4),
                    # Numeric sort keys (dropped before returning)
                    "_arch_order":   _arch_sort_val(arch),
                    "_hf_pct":       -extract_hf_pct(hf_label),  # descending
                    "_cond_left":    CONDITION_ORDER.get(short_i, 99),
                    "_cond_right":   CONDITION_ORDER.get(short_j, 99),
                })

    out = _apply_bh(pd.DataFrame(rows))
    if out.empty:
        return out

    out = (out
           .sort_values(["_arch_order", "_hf_pct", "_cond_left", "_cond_right"])
           .drop(columns=["_arch_order", "_hf_pct", "_cond_left", "_cond_right"])
           .reset_index(drop=True))
    return out

# =============================================================================
# 6. CHAMPION (POST-HOC) MCNEMAR
# =============================================================================

def run_champion_mcnemar(total_df: pd.DataFrame, top_n: int = 3) -> pd.DataFrame:
    """
    Post-hoc: within each (condition x architecture), rank models by accuracy,
    take top N, run all pairwise McNemar tests among them.

    BH correction is applied globally across all pairs in the full table.

    Sort order
    ----------
    1. Condition    : CondA → CondB → CondC
    2. Architecture : Tagalog BERT first, Tagalog DistilBERT second
    3. Accuracy of left model  : descending  (best champion first)
    4. Accuracy of right model : descending
    """
    rows = []

    for cond_key, hr_string in CONDITIONS.items():
        cond_df = total_df[
            total_df['model_key'].str.contains(hr_string, case=False)
        ]

        for arch in ['Tagalog BERT', 'Tagalog DistilBERT']:
            arch_df = cond_df[
                cond_df['model_key'].apply(extract_architecture) == arch
            ]
            if arch_df.empty:
                continue

            # Rank by accuracy
            model_acc = {
                mk: accuracy_score(grp['true_int'], grp['pred_int'])
                for mk, grp in arch_df.groupby('model_key')
            }
            top_models = sorted(model_acc, key=model_acc.get, reverse=True)[:top_n]

            if len(top_models) < 2:
                continue

            for i in range(len(top_models)):
                for j in range(i + 1, len(top_models)):
                    m1, m2 = top_models[i], top_models[j]   # m1 always higher acc

                    res_1 = arch_df[arch_df['model_key'] == m1].sort_values('sample_index')
                    res_2 = arch_df[arch_df['model_key'] == m2].sort_values('sample_index')

                    pval = _run_mcnemar_pair(res_1, res_2)
                    if pval is None:
                        continue

                    acc1 = model_acc[m1]
                    acc2 = model_acc[m2]
                    lbl1 = f"{extract_hf_label(m1)} (Acc={acc1:.4f})"
                    lbl2 = f"{extract_hf_label(m2)} (Acc={acc2:.4f})"

                    rows.append({
                        "Condition":    CONDITION_SHORT[cond_key],
                        "Architecture": arch,
                        "Comparison":   f"{lbl1} vs {lbl2}",
                        "p-value":      round(pval, 4),
                        # Numeric sort keys (dropped before returning)
                        "_cond_order":  CONDITION_ORDER.get(CONDITION_SHORT[cond_key], 99),
                        "_arch_order":  _arch_sort_val(arch),
                        "_acc_left":    -acc1,   # negative → descending
                        "_acc_right":   -acc2,
                    })

    out = _apply_bh(pd.DataFrame(rows))
    if out.empty:
        return out

    out = (out
           .sort_values(["_cond_order", "_arch_order", "_acc_left", "_acc_right"])
           .drop(columns=["_cond_order", "_arch_order", "_acc_left", "_acc_right"])
           .reset_index(drop=True))
    return out

# =============================================================================
# 7. VISUALIZATIONS
# =============================================================================

def generate_thesis_graphics(
    overall_perf_df: pd.DataFrame,
    subclass_perf_df: pd.DataFrame,
    cond_key: str,
    out_dir: str,
):
    sns.set_theme(style="white")
    cond_title = CONDITION_LABELS.get(cond_key, cond_key)

    acc_cols = [f"{sc} Accuracy" for sc in SUBCLASSES]
    h_data   = subclass_perf_df.set_index("Model Identifier")[acc_cols].copy()
    h_data.columns = [c.replace(" Accuracy", "") for c in h_data.columns]

    all_vals = pd.concat(
        [h_data.stack().reset_index(drop=True), overall_perf_df["Accuracy"]]
    )
    v_min = max(0.0, float(all_vals.min()) - 0.02)
    v_max = min(1.0, float(all_vals.max()) + 0.02)
    cmap  = "RdYlGn"
    n_m   = len(h_data)

    # ── Subclass Heatmap ──────────────────────────────────────────────────────
    fig, ax = plt.subplots(figsize=(13, max(5, n_m * 0.55 + 2.5)))
    sns.heatmap(
        h_data, ax=ax, annot=True, cmap=cmap, fmt=".4f",
        vmin=v_min, vmax=v_max,
        linewidths=0.6, linecolor="#cccccc",
        annot_kws={"size": 10, "fontname": "Arial"},
        cbar_kws={"label": "Per-Class Accuracy", "shrink": 0.75, "pad": 0.02},
    )
    ax.set_title(f"Subclass Accuracy by Model: {cond_title}",
                 fontsize=15, fontweight="bold", pad=16)
    ax.set_xlabel("News Subclass", fontsize=12, fontweight="bold", labelpad=8)
    ax.set_ylabel("Model",         fontsize=12, fontweight="bold", labelpad=8)
    ax.tick_params(axis="x", labelsize=11, rotation=0)
    ax.tick_params(axis="y", labelsize=9,  rotation=0)
    cbar = ax.collections[0].colorbar
    cbar.ax.tick_params(labelsize=9)
    cbar.set_label("Per-Class Accuracy", fontsize=10, fontweight="bold")
    plt.tight_layout()
    plt.savefig(os.path.join(out_dir, f"{cond_key}_Heatmap.png"), bbox_inches="tight")
    plt.close()

    # ── Overall Accuracy Bar Chart ────────────────────────────────────────────
    bar_data  = overall_perf_df.iloc[::-1].copy()
    norm      = plt.Normalize(vmin=v_min, vmax=v_max)
    scalar_cm = plt.cm.ScalarMappable(cmap=cmap, norm=norm)
    colors    = [scalar_cm.to_rgba(v) for v in bar_data["Accuracy"]]

    fig, ax = plt.subplots(figsize=(12, max(5, n_m * 0.6 + 2.5)))
    bars = ax.barh(
        bar_data["Model Identifier"], bar_data["Accuracy"],
        color=colors, edgecolor="#333333", linewidth=0.6, height=0.65,
    )
    for bar in bars:
        w = bar.get_width()
        ax.text(w + 0.008, bar.get_y() + bar.get_height() / 2,
                f"{w:.4f}", va="center", ha="left",
                fontsize=9, fontname="Arial", fontweight="bold", color="#222222")

    scalar_cm.set_array([])
    cbar = fig.colorbar(scalar_cm, ax=ax, orientation="vertical", shrink=0.6, pad=0.02)
    cbar.set_label("Accuracy", fontsize=10, fontweight="bold")
    cbar.ax.tick_params(labelsize=9)

    ax.set_xlabel("Overall Binary Accuracy", fontsize=12, fontweight="bold", labelpad=8)
    ax.set_title(f"Overall Model Accuracy: {cond_title}",
                 fontsize=15, fontweight="bold", pad=16)
    ax.set_xlim(0, 1.15)
    ax.xaxis.set_major_formatter(mticker.FormatStrFormatter("%.2f"))
    ax.tick_params(axis="y", labelsize=8.5)
    ax.tick_params(axis="x", labelsize=10)
    ax.spines[["top", "right"]].set_visible(False)
    plt.tight_layout()
    plt.savefig(os.path.join(out_dir, f"{cond_key}_Accuracy_Bars.png"), bbox_inches="tight")
    plt.close()

# =============================================================================
# 8. EXCEL STYLING
# =============================================================================

def apply_journal_table_styling(writer, sheet_id, df):
    """Minimalist APA-style: bold header with top+bottom rule, bottom rule on last row."""
    from openpyxl.styles import Font, Alignment, Border, Side

    ws         = writer.sheets[sheet_id]
    thin       = Side(style='thin', color="000000")
    hdr_border = Border(top=thin, bottom=thin)
    bot_border = Border(bottom=thin)
    n_cols     = len(df.columns)
    n_rows     = ws.max_row

    for c in range(1, n_cols + 1):
        cell           = ws.cell(row=1, column=c)
        cell.font      = Font(bold=True, name="Arial", size=11)
        cell.border    = hdr_border
        cell.alignment = Alignment(horizontal="center")

    for r in range(2, n_rows + 1):
        for c in range(1, n_cols + 1):
            cell           = ws.cell(row=r, column=c)
            cell.font      = Font(name="Arial", size=11)
            cell.alignment = Alignment(horizontal="center")
            cell.border    = bot_border if r == n_rows else Border()

# =============================================================================
# 9. MAIN PIPELINE
# =============================================================================

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--input_dir",  required=True)
    parser.add_argument("--output_dir", required=True)
    args = parser.parse_args()

    os.makedirs(args.output_dir, exist_ok=True)

    main_excel = os.path.join(args.output_dir, "Thesis_Results_Final_Comprehensive.xlsx")

    # ── A. Collect all master_results.xlsx recursively ───────────────────────
    print("Gathering master results...")
    frames = []
    for root, _, files in os.walk(args.input_dir):
        if "master_results.xlsx" in files and os.path.basename(root) == "results":
            path = os.path.join(root, "master_results.xlsx")
            print(f"  Found: {path}")
            frames.append(pd.read_excel(path, sheet_name="Predictions", dtype=str))

    if not frames:
        print("Error: No data found.")
        return

    total_df             = pd.concat(frames, ignore_index=True)
    total_df['true_int'] = total_df['true_label'].map(LABEL_STR_TO_INT)
    total_df['pred_int'] = total_df['pred_label'].map(LABEL_STR_TO_INT)

    main_sheets = {}

    # ── B. Per-condition metrics + within-condition McNemar ───────────────────
    for cond_key, hr_string in CONDITIONS.items():
        print(f"\nProcessing {cond_key} ({CONDITION_LABELS[cond_key]})...")
        cond_df = total_df[
            total_df['model_key'].str.contains(hr_string, case=False)
        ]
        if cond_df.empty:
            continue

        overall_rows, subclass_rows = [], []

        for (m_key, arch_raw), grp in cond_df.groupby(["model_key", "arch"]):
            yt, yp = grp['true_int'].values, grp['pred_int'].values
            prec, rec, f1, _ = precision_recall_fscore_support(
                yt, yp, average='binary', zero_division=0
            )
            overall_rows.append({
                "Model Identifier": m_key,
                "Architecture":     extract_architecture(m_key),
                "Accuracy":         round(accuracy_score(yt, yp), 4),
                "Precision":        round(prec, 4),
                "Recall":           round(rec,  4),
                "F1-Score":         round(f1,   4),
            })
            sc_row = {
                "Model Identifier": m_key,
                "Architecture":     extract_architecture(m_key),
            }
            for sc in SUBCLASSES:
                sg = grp[grp['subclass'] == sc]
                sc_row[f"{sc} Accuracy"] = (
                    round(accuracy_score(sg['true_int'], sg['pred_int']), 4)
                    if not sg.empty else 0.0
                )
            subclass_rows.append(sc_row)

        overall_df  = standardize_and_sort_thesis_data(pd.DataFrame(overall_rows))
        subclass_df = standardize_and_sort_thesis_data(pd.DataFrame(subclass_rows))
        within_df   = run_within_condition_mcnemar(cond_df)

        main_sheets[f"{cond_key}_Overall"]  = overall_df
        main_sheets[f"{cond_key}_Subclass"] = subclass_df
        main_sheets[f"{cond_key}_Stats"]    = within_df

        generate_thesis_graphics(overall_df, subclass_df, cond_key, args.output_dir)

    # ── C. Cross-condition McNemar ────────────────────────────────────────────
    print("\nRunning cross-condition McNemar...")
    main_sheets["CrossCondition_Stats"] = run_cross_condition_mcnemar(total_df)

    # ── D. Champion post-hoc McNemar ─────────────────────────────────────────
    print("\nRunning champion McNemar (top 3 per condition x architecture)...")
    main_sheets["Champion_Comparisons"] = run_champion_mcnemar(total_df, top_n=3)

    # ── E. Write single workbook (all sheets) ─────────────────────────────────
    print(f"\nWriting {main_excel}...")
    with pd.ExcelWriter(main_excel, engine='openpyxl') as writer:
        for sheet_name, df in main_sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            apply_journal_table_styling(writer, sheet_name, df)

    print("\n[COMPLETE] All results, statistics, and graphics generated.")


if __name__ == "__main__":
    main()