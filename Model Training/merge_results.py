#!/usr/bin/env python3
"""
merge_results_comprehensive_final.py
"""

import argparse
import os
import re
import warnings
from typing import Tuple

import numpy as np
import pandas as pd
import matplotlib
import matplotlib.pyplot as plt
import matplotlib.lines as mlines
import matplotlib.ticker as mticker
import seaborn as sns
from sklearn.metrics import precision_recall_fscore_support, accuracy_score

matplotlib.rcParams['figure.dpi']      = 300
matplotlib.rcParams['savefig.dpi']     = 300
matplotlib.rcParams['font.family']     = 'sans-serif'
matplotlib.rcParams['font.sans-serif'] = ['Arial']

warnings.filterwarnings("ignore")

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
CONDITION_ORDER = {v: i for i, v in enumerate(CONDITION_SHORT.values())}
ARCH_ORDER = {"Tagalog BERT": 0, "Tagalog DistilBERT": 1}

LABEL_STR_TO_INT = {"real": 0, "fake": 1}

SUBPLOT_ORDER = ["HF", "AI-F", "HR", "AI-R"]
SUBPLOT_TITLES = {
    "HF":   "HF (Human Fake)",
    "AI-F": "AI-F (AI-Generated Fake)",
    "HR":   "HR (Human Real)",
    "AI-R": "AI-R (AI-Enhanced Real)",
}

HF_TICKS    = [0, 33, 50, 67, 100]
HF_TICKLBLS = ["0%", "33%", "50%", "67%", "100%"]
HF_PCT_POS = {0: 0, 33: 1, 50: 2, 67: 3, 100: 4}
HF_POS_TICKS = list(HF_PCT_POS.values())

ARCH_STYLE = {
    "Tagalog BERT": {
        "color":  "#1f77b4",
        "marker": "o",
        "lw":     2.2,
        "ms":     7,
    },
    "Tagalog DistilBERT": {
        "color":  "#d62728",
        "marker": "s",
        "lw":     2.2,
        "ms":     7,
    },
}

COND_STYLE = {
    "CondA": {
        "label":  "Condition A (Human-Real 100%)",
        "color":  "#1f77b4",
        "marker": "o",
        "lw":     2.2,
        "ms":     7,
    },
    "CondB": {
        "label":  "Condition B (Human-Real 67%)",
        "color":  "#ff7f0e",
        "marker": "s",
        "lw":     2.2,
        "ms":     7,
    },
    "CondC": {
        "label":  "Condition C (Human-Real 50%)",
        "color":  "#2ca02c",
        "marker": "^",
        "lw":     2.2,
        "ms":     7,
    },
}

# Architecture line style for the 6-line comparison charts
ARCH_LS = {
    "Tagalog BERT":       "-",
    "Tagalog DistilBERT": "--",
}

def extract_architecture(model_key):
    return "Tagalog DistilBERT" if "distilbert" in model_key.lower() else "Tagalog BERT"

def extract_hf_label(model_key):
    key   = model_key.lower()
    hf_m  = re.search(r'hf(\d+)',  key)
    aif_m = re.search(r'aif(\d+)', key)
    hf    = int(hf_m.group(1))  if hf_m  else 0
    aif   = int(aif_m.group(1)) if aif_m else (100 - hf)
    return f"HF{hf}-AIF{aif}"

def extract_hf_pct(model_key_or_hf_label):
    m = re.search(r'hf(\d+)', model_key_or_hf_label, re.IGNORECASE)
    return int(m.group(1)) if m else 0

def format_model_display_label(model_key):
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
    hr = hr or 0; air = air or 0; hf = hf or 0; aif = aif or 0
    return f"HR{hr}-AIR{air}-HF{hf}-AIF{aif}"

def compute_logical_sort_key(model_identifier):
    mid  = model_identifier.lower()
    arch = 1 if "distilbert" in mid else 0
    hf_m = re.search(r'hf(\d+)', mid)
    hf   = int(hf_m.group(1)) if hf_m else 0
    return (arch, -hf)

def standardize_and_sort_thesis_data(df):
    if df.empty:
        return df
    if 'Model Identifier' in df.columns:
        df['_sort_key'] = df['Model Identifier'].apply(compute_logical_sort_key)
        df['Model Identifier'] = df['Model Identifier'].apply(format_model_display_label)
        df = df.sort_values('_sort_key').drop(columns=['_sort_key'])
    return df

def _run_mcnemar_pair(res_1, res_2):
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

def _apply_bh(df):
    from statsmodels.stats.multitest import multipletests
    if df.empty:
        return df
    _, p_adj, _, _ = multipletests(df['p-value'], alpha=0.05, method='fdr_bh')
    df = df.copy()
    df['p-adj (BH)']  = np.round(p_adj, 4)
    df['Significant'] = p_adj < 0.05
    return df

def _arch_sort_val(arch_str):
    return 0 if "distilbert" not in arch_str.lower() else 1

def run_within_condition_mcnemar(condition_df):
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
            if extract_hf_pct(hf_label_1) < extract_hf_pct(hf_label_2):
                hf_label_1, hf_label_2 = hf_label_2, hf_label_1
            rows.append({
                "Architecture":  extract_architecture(m1),
                "Comparison":    f"{hf_label_1} vs {hf_label_2}",
                "p-value":       round(pval, 4),
                "_arch_order":   _arch_sort_val(extract_architecture(m1)),
                "_hf_left":      -extract_hf_pct(hf_label_1),
                "_hf_right":     -extract_hf_pct(hf_label_2),
            })
    out = _apply_bh(pd.DataFrame(rows))
    if out.empty:
        return out
    out = (out.sort_values(["_arch_order", "_hf_left", "_hf_right"])
              .drop(columns=["_arch_order", "_hf_left", "_hf_right"])
              .reset_index(drop=True))
    return out

def run_cross_condition_mcnemar(total_df):
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
                    "_arch_order":   _arch_sort_val(arch),
                    "_hf_pct":       -extract_hf_pct(hf_label),
                    "_cond_left":    CONDITION_ORDER.get(short_i, 99),
                    "_cond_right":   CONDITION_ORDER.get(short_j, 99),
                })
    out = _apply_bh(pd.DataFrame(rows))
    if out.empty:
        return out
    out = (out.sort_values(["_arch_order", "_hf_pct", "_cond_left", "_cond_right"])
              .drop(columns=["_arch_order", "_hf_pct", "_cond_left", "_cond_right"])
              .reset_index(drop=True))
    return out

def run_champion_mcnemar(total_df, top_n=3):
    rows = []
    for cond_key, hr_string in CONDITIONS.items():
        cond_df = total_df[total_df['model_key'].str.contains(hr_string, case=False)]
        for arch in ['Tagalog BERT', 'Tagalog DistilBERT']:
            arch_df = cond_df[cond_df['model_key'].apply(extract_architecture) == arch]
            if arch_df.empty:
                continue
            model_acc = {
                mk: accuracy_score(grp['true_int'], grp['pred_int'])
                for mk, grp in arch_df.groupby('model_key')
            }
            top_models = sorted(model_acc, key=model_acc.get, reverse=True)[:top_n]
            if len(top_models) < 2:
                continue
            for i in range(len(top_models)):
                for j in range(i + 1, len(top_models)):
                    m1, m2 = top_models[i], top_models[j]
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
                        "_cond_order":  CONDITION_ORDER.get(CONDITION_SHORT[cond_key], 99),
                        "_arch_order":  _arch_sort_val(arch),
                        "_acc_left":    -acc1,
                        "_acc_right":   -acc2,
                    })
    out = _apply_bh(pd.DataFrame(rows))
    if out.empty:
        return out
    out = (out.sort_values(["_cond_order", "_arch_order", "_acc_left", "_acc_right"])
              .drop(columns=["_cond_order", "_arch_order", "_acc_left", "_acc_right"])
              .reset_index(drop=True))
    return out

def _build_trend_summary(total_df):
    records = []
    for model_key, grp in total_df.groupby('model_key'):
        arch    = extract_architecture(model_key)
        hf_pct  = extract_hf_pct(model_key)
        cond_key = None
        for ck, hr_str in CONDITIONS.items():
            if hr_str.lower() in model_key.lower():
                cond_key = ck
                break
        if cond_key is None:
            continue
        row = {
            'model_key':  model_key,
            'arch':       arch,
            'hf_pct':     hf_pct,
            'condition':  cond_key,
            'cond_short': CONDITION_SHORT[cond_key],
        }
        for sc in SUBCLASSES:
            sg = grp[grp['subclass'] == sc]
            row[f'{sc}_acc'] = (
                accuracy_score(sg['true_int'], sg['pred_int']) * 100
                if not sg.empty else np.nan
            )
        records.append(row)
    return pd.DataFrame(records).sort_values(['condition', 'arch', 'hf_pct'])

def _style_trend_axes(ax, title):
    ax.set_title(title, fontsize=12, fontweight='bold', pad=6)
    ax.set_xlabel('HF Portion →', fontsize=10, labelpad=4)
    ax.set_ylabel('Acc (%)', fontsize=10, labelpad=4)
    ax.set_xticks(HF_POS_TICKS)
    ax.set_xticklabels(HF_TICKLBLS, fontsize=9)
    ax.tick_params(axis='y', labelsize=9)
    ax.grid(True, color='#e0e0e0', linewidth=0.8, zorder=0)
    ax.set_axisbelow(True)
    ax.spines[['top', 'right']].set_visible(False)
    ax.set_ylim(bottom=0, top=105)   # give room for 100% lines


# =============================================================================
# 7a. FIGURE 4 STYLE — per condition, both architectures as lines (3 charts)
# =============================================================================

def generate_per_condition_trend_charts(summary_df, out_dir):
    FIG_SIZE = (11, 8.5)

    for cond_key, cond_label in CONDITION_LABELS.items():
        cond_df = summary_df[summary_df['condition'] == cond_key]
        if cond_df.empty:
            continue

        fig, axes = plt.subplots(
            2, 2, figsize=FIG_SIZE,
            gridspec_kw={'hspace': 0.45, 'wspace': 0.28}
        )
        fig.subplots_adjust(top=0.80)            # pull subplots down a bit

        ax_flat = axes.flatten()
        for idx, sc in enumerate(SUBPLOT_ORDER):
            ax = ax_flat[idx]
            for arch, style in ARCH_STYLE.items():
                sub = cond_df[cond_df['arch'] == arch].sort_values('hf_pct')
                if sub.empty:
                    continue
                x_vals = sub['hf_pct'].map(HF_PCT_POS)
                ax.plot(
                    x_vals, sub[f'{sc}_acc'],
                    color=style['color'], marker=style['marker'],
                    linewidth=style['lw'], markersize=style['ms'],
                    label=arch, zorder=3,
                )
            _style_trend_axes(ax, SUBPLOT_TITLES[sc])

        handles = [
            mlines.Line2D([], [], color=style['color'], marker=style['marker'],
                          linewidth=style['lw'], markersize=style['ms'], label=arch)
            for arch, style in ARCH_STYLE.items()
        ]
        fig.legend(
            handles=handles,
            loc='upper center',
            bbox_to_anchor=(0.5, 0.94),
            ncol=2,
            fontsize=10,
            framealpha=0.9,
            edgecolor='#cccccc',
            handlelength=2.5,
        )
        fig.suptitle(
            f'Subclass Accuracy Across Training Ratios\n{cond_label}',
            fontsize=14, fontweight='bold', y=0.885,   # raised from 0.88
        )

        fname = os.path.join(out_dir, f'{cond_key}_RatioTrend.png')
        fig.savefig(fname, bbox_inches='tight', dpi=300)
        plt.close(fig)
        print(f'  Saved: {fname}')

# 7b. FIGURE 5 STYLE — per condition, both architectures + all conditions (3 charts)
# =============================================================================

def generate_per_arch_condition_comparison_charts(summary_df, out_dir):
    FIG_SIZE = (11, 8.5)

    for arch in ['Tagalog BERT', 'Tagalog DistilBERT']:
        arch_df = summary_df[summary_df['arch'] == arch]
        if arch_df.empty:
            continue

        fig, axes = plt.subplots(
            2, 2, figsize=FIG_SIZE,
            gridspec_kw={'hspace': 0.45, 'wspace': 0.28}
        )
        fig.subplots_adjust(top=0.80)            # pull subplots down a bit

        ax_flat = axes.flatten()
        for idx, sc in enumerate(SUBPLOT_ORDER):
            ax = ax_flat[idx]
            for ck in CONDITIONS:
                short = CONDITION_SHORT[ck]
                style = COND_STYLE[short]
                sub = arch_df[arch_df['condition'] == ck].sort_values('hf_pct')
                if sub.empty:
                    continue
                x_vals = sub['hf_pct'].map(HF_PCT_POS)
                ax.plot(
                    x_vals, sub[f'{sc}_acc'],
                    color=style['color'],
                    linestyle='-',
                    marker=style['marker'],
                    linewidth=style['lw'],
                    markersize=style['ms'],
                    alpha=1.0,
                    label=CONDITION_LABELS[ck],
                    zorder=3,
                )
            _style_trend_axes(ax, SUBPLOT_TITLES[sc])

        handles = [
            mlines.Line2D([], [], color=COND_STYLE[CONDITION_SHORT[ck]]['color'],
                          linestyle='-',
                          marker=COND_STYLE[CONDITION_SHORT[ck]]['marker'],
                          linewidth=2.2, markersize=7,
                          label=CONDITION_LABELS[ck])
            for ck in CONDITIONS
        ]
        fig.legend(
            handles=handles,
            loc='upper center',
            bbox_to_anchor=(0.5, 0.94),
            ncol=3,
            fontsize=9.5,
            framealpha=0.9,
            edgecolor='#cccccc',
            handlelength=2.5,
        )
        fig.suptitle(
            f'Subclass Accuracy Across Training Ratios\n{arch}',
            fontsize=14, fontweight='bold', y=0.885,   # raised from 0.88
        )

        safe_arch = arch.replace(' ', '_')
        fname = os.path.join(out_dir, f'{safe_arch}_ConditionComparison_RatioTrend.png')
        fig.savefig(fname, bbox_inches='tight', dpi=300)
        plt.close(fig)
        print(f'  Saved: {fname}')

def generate_overall_accuracy_bars(overall_df, cond_key, out_dir):
    """
    Grouped horizontal accuracy bars: one group per model setting,
    two bars (BERT vs DistilBERT) offset vertically.
    Architecture displayed in a legend. Y-axis shows only Model Identifier.
    Label order matches the original chart (highest accuracy on top).
    """
    cond_title = CONDITION_LABELS.get(cond_key, cond_key)

    # Replicate the original ordering: reverse the full dataframe to put
    # DistilBERT first, then BERT, both sorted by HF% descending.
    # (This matches the iloc[::-1] that was used before.)
    bar_data = overall_df.iloc[::-1].copy()

    # Colors from ARCH_STYLE
    arch_color = {
        "Tagalog BERT":       ARCH_STYLE["Tagalog BERT"]["color"],
        "Tagalog DistilBERT": ARCH_STYLE["Tagalog DistilBERT"]["color"],
    }

    # Get unique model identifiers in the order they appear in bar_data
    model_order = bar_data["Model Identifier"].drop_duplicates().tolist()

    # Build an accuracy lookup: (model_id, arch) -> accuracy
    acc_lookup = {}
    for _, row in bar_data.iterrows():
        acc_lookup[(row["Model Identifier"], row["Architecture"])] = row["Accuracy"]

    n_groups = len(model_order)
    y = np.arange(n_groups)                # centre of each group
    bar_height = 0.25                      # half‑height for each side
    fig, ax = plt.subplots(figsize=(14, max(6, n_groups * 1.2 + 2)))

    # Draw BERT bars slightly above the centre, DistilBERT slightly below
    for arch, offset in [("Tagalog BERT", +bar_height/2),
                         ("Tagalog DistilBERT", -bar_height/2)]:
        vals = [acc_lookup.get((model, arch), 0) for model in model_order]
        bars = ax.barh(y + offset, vals, height=bar_height,
                       color=arch_color[arch],
                       edgecolor="#333333",
                       linewidth=0.6,
                       label=arch)
        # Value labels
        for bar in bars:
            w = bar.get_width()
            ax.text(w + 0.008, bar.get_y() + bar.get_height()/2,
                    f"{w:.4f}", va="center", ha="left",
                    fontsize=9, fontweight="bold", color="#222222")

    ax.set_yticks(y)
    ax.set_yticklabels(model_order, fontsize=9)
    ax.set_xlabel("Overall Binary Accuracy", fontsize=12, fontweight="bold", labelpad=8)
    ax.set_title(f"Overall Model Accuracy: {cond_title}", fontsize=15, fontweight="bold", pad=16)
    ax.set_xlim(0, 1.15)
    ax.xaxis.set_major_formatter(mticker.FormatStrFormatter("%.2f"))
    ax.legend(loc="lower right", framealpha=0.9)
    ax.spines[["top", "right"]].set_visible(False)
    ax.invert_yaxis()   # keep highest accuracy at the top
    plt.tight_layout()
    plt.savefig(os.path.join(out_dir, f"{cond_key}_Accuracy_Bars.png"), bbox_inches="tight")
    plt.close()

def apply_journal_table_styling(writer, sheet_id, df):
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

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--input_dir",  required=True)
    parser.add_argument("--output_dir", required=True)
    args = parser.parse_args()

    os.makedirs(args.output_dir, exist_ok=True)
    main_excel = os.path.join(args.output_dir, "Thesis_Results_Final_Comprehensive.xlsx")

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
    overall_dfs = {}

    for cond_key, hr_string in CONDITIONS.items():
        print(f"\nProcessing {cond_key} ({CONDITION_LABELS[cond_key]})...")
        cond_df = total_df[total_df['model_key'].str.contains(hr_string, case=False)]
        if cond_df.empty:
            continue

        overall_rows, subclass_rows = [], []
        for (m_key, arch_raw), grp in cond_df.groupby(["model_key", "arch"]):
            yt, yp = grp['true_int'].values, grp['pred_int'].values
            prec, rec, f1, _ = precision_recall_fscore_support(
                yt, yp, average='binary', zero_division=0)
            overall_rows.append({
                "Model Identifier": m_key,
                "Architecture":     extract_architecture(m_key),
                "Accuracy":         round(accuracy_score(yt, yp), 4),
                "Precision":        round(prec, 4),
                "Recall":           round(rec,  4),
                "F1-Score":         round(f1,   4),
            })
            sc_row = {"Model Identifier": m_key, "Architecture": extract_architecture(m_key)}
            for sc in SUBCLASSES:
                sg = grp[grp['subclass'] == sc]
                sc_row[f"{sc} Accuracy"] = (
                    round(accuracy_score(sg['true_int'], sg['pred_int']), 4)
                    if not sg.empty else 0.0)
            subclass_rows.append(sc_row)

        overall_df  = standardize_and_sort_thesis_data(pd.DataFrame(overall_rows))
        subclass_df = standardize_and_sort_thesis_data(pd.DataFrame(subclass_rows))
        within_df   = run_within_condition_mcnemar(cond_df)

        main_sheets[f"{cond_key}_Overall"]  = overall_df
        main_sheets[f"{cond_key}_Subclass"] = subclass_df
        main_sheets[f"{cond_key}_Stats"]    = within_df
        overall_dfs[cond_key]               = overall_df

        generate_overall_accuracy_bars(overall_df, cond_key, args.output_dir)

    print("\nGenerating ratio trend charts...")
    trend_summary = _build_trend_summary(total_df)

    print("  [Fig-4 style] Per-condition trend charts (both archs, legend above)...")
    generate_per_condition_trend_charts(trend_summary, args.output_dir)

    print("  [Comparison charts] One figure per architecture (3 condition lines each)...")
    generate_per_arch_condition_comparison_charts(trend_summary, args.output_dir)

    print("\nRunning cross-condition McNemar...")
    main_sheets["CrossCondition_Stats"] = run_cross_condition_mcnemar(total_df)

    print("\nRunning champion McNemar (top 3 per condition x architecture)...")
    main_sheets["Champion_Comparisons"] = run_champion_mcnemar(total_df, top_n=3)

    print(f"\nWriting {main_excel}...")
    with pd.ExcelWriter(main_excel, engine='openpyxl') as writer:
        for sheet_name, df in main_sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            apply_journal_table_styling(writer, sheet_name, df)

    print("\n" + "=" * 60)
    print("OUTPUT SUMMARY")
    print("=" * 60)
    print(f"\n[Excel]  {main_excel}")
    for sheet in main_sheets:
        print(f"         \u2514\u2500 {sheet}")
    print("\n[Charts] Ratio Trend (per condition, both archs):")
    for ck in CONDITIONS:
        print(f"         \u2514\u2500 {ck}_RatioTrend.png")
    print("\n[Charts] Architecture Comparison (2 charts, 3 condition lines each):")
    for arch in ["Tagalog_BERT", "Tagalog_DistilBERT"]:
        print(f"         \u2514\u2500 {arch}_ConditionComparison_RatioTrend.png")
    print("\n[Charts] Overall Accuracy Bars (per condition):")
    for ck in CONDITIONS:
        print(f"         \u2514\u2500 {ck}_Accuracy_Bars.png")
    print("\n[COMPLETE] All results, statistics, and graphics generated.")


if __name__ == "__main__":
    main()