#!/usr/bin/env python3
"""
merge_results_comprehensive_final.py
=============================================================================
THE COMPLETE UNABRIDGED THESIS ANALYTICS ENGINE
=============================================================================
This script provides the full statistical and visual breakdown for 
Filipino Fake News Detection models.

CONTAINS:
- Full Binary Metrics (Precision, Recall, F1, Accuracy) per model.
- Subclass Accuracy breakdown (HR, AI-R, HF, AI-F) per model.
- McNemar Pairwise Statistical Significance with BH Correction.
- Academic Minimalist Table Styling (Arial).
- RdYlGn (Red-to-Green) Synchronized Visualizations.
"""

import argparse
import os
import re
import warnings
from typing import Dict, List, Tuple

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import seaborn as sns
from sklearn.metrics import precision_recall_fscore_support, accuracy_score

# Set global plotting parameters for academic publication
plt.rcParams['figure.dpi'] = 300
plt.rcParams['savefig.dpi'] = 300
plt.rcParams['font.family'] = 'sans-serif'
plt.rcParams['font.sans-serif'] = ['Arial']

# Suppress warnings to keep thesis logs clean
warnings.filterwarnings("ignore")

# =============================================================================
# 1. GLOBAL CONSTANTS
# =============================================================================

SUBCLASSES = ["HR", "AI-R", "HF", "AI-F"]
CONDITIONS = {
    "Condition_A": "HR100",
    "Condition_B": "HR67",
    "Condition_C": "HR50"
}
CONDITION_LABELS = {
    "Condition_A": "Condition A: Human-Real 100%",
    "Condition_B": "Condition B: Human-Real 67%",
    "Condition_C": "Condition C: Human-Real 50%",
}
LABEL_STR_TO_INT = {"real": 0, "fake": 1}

# =============================================================================
# 2. LOGICAL SORTING ENGINE
# =============================================================================

def compute_logical_sort_key(model_identifier: str) -> Tuple:
    """
    Ensures models are grouped by architecture (BERT then DistilBERT)
    and then by training data composition (HF Ratio descending).
    """
    mid_lower = str(model_identifier).lower()
    
    # 1. Architecture priority
    if "distilbert" in mid_lower:
        architecture_weight = 1
    else:
        architecture_weight = 0
        
    # 2. Training data composition priority (HF100 > HF75 > HF50...)
    hf_match = re.search(r'hf(\d+)', mid_lower)
    if hf_match:
        hf_composition_ratio = int(hf_match.group(1))
    else:
        hf_composition_ratio = 0
        
    return (architecture_weight, -hf_composition_ratio)


def format_model_display_label(model_key: str) -> str:
    """
    Converts a raw model key into a human-readable publication label.

    Examples
    --------
    "bert-HR67-HF100" -> "Tagalog BERT: HR67-AIR33-HF100-AIF0"
    "distilbert-HR50-HF75" -> "Tagalog DistilBERT: HR50-AIR50-HF75-AIF25"

    The function infers missing ratio values so that all four components
    (HR, AIR, HF, AIF) always sum correctly.
    """
    key_lower = model_key.lower()

    # --- Architecture ---
    if "distilbert" in key_lower:
        arch_label = "Tagalog DistilBERT"
    else:
        arch_label = "Tagalog BERT"

    # --- Extract numeric ratios (default to None if absent) ---
    def _extract(pattern, text):
        m = re.search(pattern, text, re.IGNORECASE)
        return int(m.group(1)) if m else None

    hr  = _extract(r'hr(\d+)',  key_lower)
    air = _extract(r'air(\d+)', key_lower)
    hf  = _extract(r'hf(\d+)',  key_lower)
    aif = _extract(r'aif(\d+)', key_lower)

    # Infer missing complements (assumes pairs sum to 100)
    if hr is not None and air is None:
        air = 100 - hr
    elif air is not None and hr is None:
        hr = 100 - air

    if hf is not None and aif is None:
        aif = 100 - hf
    elif aif is not None and hf is None:
        hf = 100 - aif

    # Fallback to 0 if still unresolved
    hr  = hr  if hr  is not None else 0
    air = air if air is not None else 0
    hf  = hf  if hf  is not None else 0
    aif = aif if aif is not None else 0

    composition = f"HR{hr}-AIR{air}-HF{hf}-AIF{aif}"
    return f"{arch_label}: {composition}"


def standardize_and_sort_thesis_data(dataframe: pd.DataFrame) -> pd.DataFrame:
    """Standardizes architecture labels and applies the logical thesis sort."""
    if dataframe.empty:
        return dataframe
        
    # Standardize Architecture labels (no hyphens)
    if 'Architecture' in dataframe.columns:
        dataframe['Architecture'] = dataframe['Architecture'].replace({
            'BERT':       'Tagalog BERT',
            'DISTILBERT': 'Tagalog DistilBERT',
            'bert':       'Tagalog BERT',
            'distilbert': 'Tagalog DistilBERT',
            # also handle legacy hyphenated variants
            'Tagalog-BERT':       'Tagalog BERT',
            'Tagalog-DistilBERT': 'Tagalog DistilBERT',
        })

    # Apply formatted display labels to Model Identifier
    if 'Model Identifier' in dataframe.columns:
        dataframe['Model Identifier'] = dataframe['Model Identifier'].apply(
            format_model_display_label
        )

    # Sort the dataframe using our logic (operates on the new label string)
    if 'Model Identifier' in dataframe.columns:
        dataframe['_sort_key'] = dataframe['Model Identifier'].apply(compute_logical_sort_key)
        dataframe = dataframe.sort_values(by='_sort_key').drop(columns=['_sort_key'])
        
    return dataframe

# =============================================================================
# 3. STATISTICAL ENGINE (MCNEMAR + BH CORRECTION)
# =============================================================================

def run_mcnemar_significance_test(condition_dataframe: pd.DataFrame) -> pd.DataFrame:
    """
    Performs pairwise McNemar tests to evaluate if model performance 
    differences are statistically significant.
    """
    from statsmodels.stats.multitest import multipletests
    from statsmodels.stats.contingency_tables import mcnemar
    
    unique_model_list = condition_dataframe['model_key'].unique()
    pairwise_stats_list = []
    
    for i in range(len(unique_model_list)):
        for j in range(i + 1, len(unique_model_list)):
            model_1_name = unique_model_list[i]
            model_2_name = unique_model_list[j]
            
            # Synchronize results by sample index
            res_1 = condition_dataframe[condition_dataframe['model_key'] == model_1_name].sort_values('sample_index')
            res_2 = condition_dataframe[condition_dataframe['model_key'] == model_2_name].sort_values('sample_index')
            
            common_count = min(len(res_1), len(res_2))
            ground_truth = res_1['true_int'].values[:common_count]
            pred_m1 = res_1['pred_int'].values[:common_count]
            pred_m2 = res_2['pred_int'].values[:common_count]
            
            # Contingency table values
            b_count = np.sum((pred_m1 == ground_truth) & (pred_m2 != ground_truth))
            c_count = np.sum((pred_m1 != ground_truth) & (pred_m2 == ground_truth))
            
            contingency_matrix = [[0, b_count], [c_count, 0]]
            mcnemar_result = mcnemar(contingency_matrix, exact=(b_count + c_count < 25))
            
            pairwise_stats_list.append({
                "Comparison": f"{format_model_display_label(model_1_name)} vs {format_model_display_label(model_2_name)}",
                "Discordant_M1": b_count,
                "Discordant_M2": c_count,
                "p-value": mcnemar_result.pvalue
            })
            
    if not pairwise_stats_list:
        return pd.DataFrame()
        
    stat_df = pd.DataFrame(pairwise_stats_list)
    
    _, p_adj, _, _ = multipletests(stat_df['p-value'], alpha=0.05, method='fdr_bh')
    stat_df['p-adj (BH)'] = p_adj
    stat_df['Significant'] = p_adj < 0.05
    
    return stat_df

# =============================================================================
# 4. ACADEMIC VISUALIZATION (HEATMAPS & RDYLGN BARS)
# =============================================================================

def generate_thesis_graphics(
    overall_perf_df: pd.DataFrame,
    subclass_perf_df: pd.DataFrame,
    cond_key: str,
    out_dir: str,
):
    """
    Generates publication-quality figures.

    Heatmap improvements
    --------------------
    * vmin / vmax are derived from the actual data range so the full
      RdYlGn spectrum is used rather than being anchored at an arbitrary
      0.85 centre — every shade of the palette carries meaning.
    * A descriptive colorbar label is included.
    * Minor grid lines are suppressed; cell borders are kept thin.

    Bar chart improvements
    ----------------------
    * Colour is normalised over the same [vmin, vmax] window as the
      heatmap so the two figures are visually consistent.
    * A shared colorbar legend is attached to the bar chart.
    * Axis labels and annotation font sizes are harmonised.
    """
    sns.set_theme(style="white")

    cond_title = CONDITION_LABELS.get(cond_key, cond_key)

    # ------------------------------------------------------------------ #
    # Shared colour scale — computed once and reused in both figures       #
    # ------------------------------------------------------------------ #
    acc_cols = [f"{sc} Accuracy" for sc in SUBCLASSES]
    h_data   = subclass_perf_df.set_index("Model Identifier")[acc_cols].copy()
    h_data.columns = [c.replace(" Accuracy", "") for c in h_data.columns]

    all_acc_vals = pd.concat(
        [h_data.stack().reset_index(drop=True),
         overall_perf_df["Accuracy"]]
    )
    v_min = max(0.0, float(all_acc_vals.min()) - 0.02)   # small breathing room
    v_max = min(1.0, float(all_acc_vals.max()) + 0.02)

    cmap = "RdYlGn"

    # ------------------------------------------------------------------ #
    # A. Subclass Accuracy Heatmap                                         #
    # ------------------------------------------------------------------ #
    n_models = len(h_data)
    fig_h    = max(5, n_models * 0.55 + 2.5)
    fig, ax  = plt.subplots(figsize=(13, fig_h))

    sns.heatmap(
        h_data,
        ax=ax,
        annot=True,
        cmap=cmap,
        fmt=".4f",
        vmin=v_min,
        vmax=v_max,
        linewidths=0.6,
        linecolor="#cccccc",
        annot_kws={"size": 10, "fontname": "Arial"},
        cbar_kws={"label": "Per-Class Accuracy", "shrink": 0.75, "pad": 0.02},
    )

    ax.set_title(
        f"Subclass Accuracy by Model: {cond_title}",
        fontsize=15,
        fontweight="bold",
        pad=16,
    )
    ax.set_xlabel("News Subclass", fontsize=12, fontweight="bold", labelpad=8)
    ax.set_ylabel("Model", fontsize=12, fontweight="bold", labelpad=8)
    ax.tick_params(axis="x", labelsize=11, rotation=0)
    ax.tick_params(axis="y", labelsize=9,  rotation=0)

    # Colorbar font
    cbar = ax.collections[0].colorbar
    cbar.ax.tick_params(labelsize=9)
    cbar.set_label("Per-Class Accuracy", fontsize=10, fontweight="bold")

    plt.tight_layout()
    plt.savefig(os.path.join(out_dir, f"{cond_key}_Heatmap.png"), bbox_inches="tight")
    plt.close()

    # ------------------------------------------------------------------ #
    # B. Overall Accuracy Horizontal Bar Chart                             #
    # ------------------------------------------------------------------ #
    bar_data = overall_perf_df.iloc[::-1].copy()   # invert for top-to-bottom reading

    norm      = plt.Normalize(vmin=v_min, vmax=v_max)
    scalar_cm = plt.cm.ScalarMappable(cmap=cmap, norm=norm)
    bar_colors = [scalar_cm.to_rgba(v) for v in bar_data["Accuracy"]]

    fig, ax = plt.subplots(figsize=(12, max(5, n_models * 0.6 + 2.5)))

    bars = ax.barh(
        bar_data["Model Identifier"],
        bar_data["Accuracy"],
        color=bar_colors,
        edgecolor="#333333",
        linewidth=0.6,
        height=0.65,
    )

    # Value annotations
    for bar in bars:
        width = bar.get_width()
        ax.text(
            width + 0.008,
            bar.get_y() + bar.get_height() / 2,
            f"{width:.4f}",
            va="center",
            ha="left",
            fontsize=9,
            fontname="Arial",
            fontweight="bold",
            color="#222222",
        )

    # Colorbar legend (matches heatmap scale)
    scalar_cm.set_array([])
    cbar = fig.colorbar(scalar_cm, ax=ax, orientation="vertical",
                        shrink=0.6, pad=0.02)
    cbar.set_label("Accuracy", fontsize=10, fontweight="bold")
    cbar.ax.tick_params(labelsize=9)

    ax.set_xlabel("Overall Binary Accuracy", fontsize=12, fontweight="bold", labelpad=8)
    ax.set_title(
        f"Overall Model Accuracy: {cond_title}",
        fontsize=15,
        fontweight="bold",
        pad=16,
    )
    ax.set_xlim(0, 1.15)
    ax.xaxis.set_major_formatter(mticker.FormatStrFormatter("%.2f"))
    ax.tick_params(axis="y", labelsize=8.5)
    ax.tick_params(axis="x", labelsize=10)
    ax.spines[["top", "right"]].set_visible(False)

    plt.tight_layout()
    plt.savefig(os.path.join(out_dir, f"{cond_key}_Accuracy_Bars.png"), bbox_inches="tight")
    plt.close()

# =============================================================================
# 5. JOURNAL-STYLE EXCEL STYLING
# =============================================================================

def apply_journal_table_styling(excel_writer, sheet_id, dataframe):
    """Applies minimalist APA-style horizontal formatting."""
    from openpyxl.styles import Font, Alignment, Border, Side
    worksheet = excel_writer.sheets[sheet_id]
    
    black_thin_side = Side(style='thin', color="000000")
    header_border_style = Border(top=black_thin_side, bottom=black_thin_side)
    bottom_border_style = Border(bottom=black_thin_side)

    for c_idx in range(1, len(dataframe.columns) + 1):
        cell_obj = worksheet.cell(row=1, column=c_idx)
        cell_obj.font = Font(bold=True, name="Arial", size=11)
        cell_obj.border = header_border_style
        cell_obj.alignment = Alignment(horizontal="center")

    total_rows = worksheet.max_row
    for r_idx in range(2, total_rows + 1):
        for c_idx in range(1, len(dataframe.columns) + 1):
            cell_obj = worksheet.cell(row=r_idx, column=c_idx)
            cell_obj.font = Font(name="Arial", size=11)
            cell_obj.alignment = Alignment(horizontal="center")
            if r_idx == total_rows:
                cell_obj.border = bottom_border_style
            else:
                cell_obj.border = Border()

# =============================================================================
# 6. MAIN PIPELINE
# =============================================================================

def main():
    arg_parser = argparse.ArgumentParser()
    arg_parser.add_argument("--input_dir", required=True)
    arg_parser.add_argument("--output_dir", required=True)
    args = arg_parser.parse_args()

    os.makedirs(args.output_dir, exist_ok=True)
    excel_report_path = os.path.join(args.output_dir, "Thesis_Results_Final_Comprehensive.xlsx")

    # A. Aggregate Data
    print("Gathering master results from experimental directories...")
    master_pred_list = []
    for item in os.listdir(args.input_dir):
        potential_path = os.path.join(args.input_dir, item, "master_results.xlsx")
        if os.path.exists(potential_path):
            pred_df = pd.read_excel(potential_path, sheet_name="Predictions", dtype=str)
            master_pred_list.append(pred_df)
    
    if not master_pred_list:
        print("Error: No data found."); return
        
    total_data_df = pd.concat(master_pred_list, ignore_index=True)
    total_data_df['true_int'] = total_data_df['true_label'].map(LABEL_STR_TO_INT)
    total_data_df['pred_int'] = total_data_df['pred_label'].map(LABEL_STR_TO_INT)

    all_thesis_sheets = {}

    # B. Generate Results by Condition
    for cond_key, hr_string in CONDITIONS.items():
        print(f"Processing {cond_key} ({CONDITION_LABELS[cond_key]})...")
        condition_subset = total_data_df[total_data_df['model_key'].str.contains(hr_string, case=False)]
        
        if condition_subset.empty:
            continue
            
        overall_rows = []
        subclass_rows = []
        
        for (m_key, arch_raw), group in condition_subset.groupby(["model_key", "arch"]):
            y_t, y_p = group['true_int'].values, group['pred_int'].values
            
            prec_val, recall_val, f1_val, _ = precision_recall_fscore_support(
                y_t, y_p, average='binary', zero_division=0
            )
            overall_rows.append({
                "Model Identifier": m_key,
                "Architecture": arch_raw,
                "Accuracy":  round(accuracy_score(y_t, y_p), 4),
                "Precision": round(prec_val,   4),
                "Recall":    round(recall_val, 4),
                "F1-Score":  round(f1_val,     4),
            })
            
            sc_meta = {"Model Identifier": m_key, "Architecture": arch_raw}
            for sc_name in SUBCLASSES:
                sc_group = group[group['subclass'] == sc_name]
                if not sc_group.empty:
                    sc_meta[f"{sc_name} Accuracy"] = round(
                        accuracy_score(sc_group['true_int'], sc_group['pred_int']), 4
                    )
                else:
                    sc_meta[f"{sc_name} Accuracy"] = 0.0
            subclass_rows.append(sc_meta)

        overall_perf_df  = standardize_and_sort_thesis_data(pd.DataFrame(overall_rows))
        subclass_perf_df = standardize_and_sort_thesis_data(pd.DataFrame(subclass_rows))
        significance_df  = run_mcnemar_significance_test(condition_subset)

        all_thesis_sheets[f"{cond_key}_Overall"]  = overall_perf_df
        all_thesis_sheets[f"{cond_key}_Subclass"] = subclass_perf_df
        all_thesis_sheets[f"{cond_key}_Stats"]    = significance_df
        
        generate_thesis_graphics(overall_perf_df, subclass_perf_df, cond_key, args.output_dir)

    # C. Final Excel Export
    print(f"\nWriting reports to {excel_report_path}...")
    with pd.ExcelWriter(excel_report_path, engine='openpyxl') as writer:
        for sheet_name, df_content in all_thesis_sheets.items():
            df_content.to_excel(writer, sheet_name=sheet_name, index=False)
            apply_journal_table_styling(writer, sheet_name, df_content)

    print("\n[COMPLETE] All academic results, stats, and synchronized graphics generated.")


if __name__ == "__main__":
    main()