#!/usr/bin/env python3
"""
merge_results_comprehensive_thesis.py
==============================================
The DEFINITIVE Thesis Results Engine.

FEATURES:
1. DATA SEGMENTATION: Auto-splits results into Condition A, B, and C.
2. HORIZONTAL METRICS: Generates P/R/F1 and Subclass Accuracy in wide format.
3. STATISTICAL RIGOR: Full McNemar Test + Benjamini-Hochberg FDR Correction.
4. FIXED-ORDER VISUALS: 
   - Bar charts follow logical order (Arch -> HF Ratio), NOT result value.
   - Heatmaps use 'RdYlGn' (Red-Yellow-Green) traffic light palette.
5. PROFESSIONAL STYLING: Navy/Gold/Blue Excel themes for publication.
"""

import argparse
import os
import re
import warnings
from typing import Dict, List, Tuple

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from sklearn.metrics import precision_recall_fscore_support, accuracy_score

# Ensure high-quality plots for thesis printing
plt.rcParams['figure.dpi'] = 300
plt.rcParams['savefig.dpi'] = 300
warnings.filterwarnings("ignore")

# =============================================================================
# GLOBAL CONFIGURATION
# =============================================================================

SUBCLASSES = ["HR", "AI-R", "HF", "AI-F"]
CONDITIONS = {
    "Condition_A": "HR100",
    "Condition_B": "HR67",
    "Condition_C": "HR50"
}
LABEL_MAP = {"real": 0, "fake": 1}

# Excel Professional Palette
_HDR_BG = "1F4E79"  # Deep Navy
_HDR_FG = "FFFFFF"  # White
_ALT_ROW = "D6E4F0" # Light Blue
_SIG_BG = "FFEB9C"  # Gold Significance
_BORDER = "BDD7EE"  # Border Blue

# =============================================================================
# LOGICAL SORTING ENGINE
# =============================================================================

def get_logical_sort_key(model_key: str) -> Tuple:
    """
    Ensures consistent order:
    1. Tagalog-BERT comes before Tagalog-DistilBERT.
    2. Higher HF ratios (100) come before lower HF ratios (0).
    """
    k = str(model_key).lower()
    # Architecture weight
    arch_weight = 1 if "distilbert" in k else 0
    # HF Ratio extraction
    hf_match = re.search(r'hf(\d+)', k)
    hf_val = int(hf_match.group(1)) if hf_match else 0
    
    return (arch_weight, -hf_val)

def standardize_and_sort(df: pd.DataFrame) -> pd.DataFrame:
    """Applies thesis-standard names and logical sorting."""
    if df.empty: return df
    
    if 'Architecture' in df.columns:
        df['Architecture'] = df['Architecture'].replace({
            'BERT': 'Tagalog-BERT', 
            'DISTILBERT': 'Tagalog-DistilBERT'
        })
        
    if 'Model Identifier' in df.columns:
        df['_sort_val'] = df['Model Identifier'].apply(get_logical_sort_key)
        df = df.sort_values(by='_sort_val').drop(columns=['_sort_val'])
        
    return df

# =============================================================================
# STATISTICAL ENGINE (MCNEMAR + BH)
# =============================================================================

def calculate_mcnemar_stats(condition_df: pd.DataFrame) -> pd.DataFrame:
    """Performs pairwise McNemar tests with FDR correction for the condition."""
    from statsmodels.stats.multitest import multipletests
    from statsmodels.stats.contingency_tables import mcnemar
    
    models = condition_df['model_key'].unique()
    results = []
    
    for i in range(len(models)):
        for j in range(i + 1, len(models)):
            m1, m2 = models[i], models[j]
            
            # Align predictions by sample index
            d1 = condition_df[condition_df['model_key'] == m1].sort_values('sample_index')
            d2 = condition_df[condition_df['model_key'] == m2].sort_values('sample_index')
            
            common_len = min(len(d1), len(d2))
            y_true = d1['true_label'].values[:common_len]
            p1 = d1['pred_label'].values[:common_len]
            p2 = d2['pred_label'].values[:common_len]
            
            # Contingency matrix elements
            # b: M1 Correct, M2 Wrong
            # c: M1 Wrong, M2 Correct
            b = np.sum((p1 == y_true) & (p2 != y_true))
            c = np.sum((p1 != y_true) & (p2 == y_true))
            
            # McNemar Test
            table = [[0, b], [c, 0]]
            # Use exact binomial if discordant count is small
            test_res = mcnemar(table, exact=(b + c < 25))
            
            results.append({
                "Comparison": f"{m1} vs {m2}",
                "Discordant_M1": b,
                "Discordant_M2": c,
                "p-value": test_res.pvalue
            })
            
    if not results: return pd.DataFrame()
    
    stat_df = pd.DataFrame(results)
    # Apply Benjamini-Hochberg Correction
    _, p_adj, _, _ = multipletests(stat_df['p-value'], alpha=0.05, method='fdr_bh')
    stat_df['p-adj (BH)'] = p_adj
    stat_df['Significant'] = p_adj < 0.05
    
    return stat_df

# =============================================================================
# VISUALIZATION ENGINE (RED-TO-GREEN & FIXED ORDER)
# =============================================================================

def create_plots(overall_df: pd.DataFrame, subclass_df: pd.DataFrame, cond_name: str, out_dir: str):
    """Generates charts that respect the logical order, not the result values."""
    sns.set_theme(style="whitegrid")
    
    # 1. HEATMAP (Red -> Yellow -> Green)
    plt.figure(figsize=(12, 8))
    # Extract subclass columns
    cols = [f"{sc} Accuracy" for sc in SUBCLASSES]
    h_data = subclass_df.set_index("Model Identifier")[cols]
    h_data.columns = [c.replace(" Accuracy", "") for c in h_data.columns]
    
    # 'RdYlGn' is the Red-Yellow-Green colormap
    sns.heatmap(h_data, annot=True, cmap="RdYlGn", fmt=".4f", linewidths=.5, cbar_kws={'label': 'Accuracy'})
    plt.title(f"Subclass Accuracy Heatmap: {cond_name.replace('_', ' ')}", fontsize=15, pad=20)
    plt.tight_layout()
    plt.savefig(os.path.join(out_dir, f"{cond_name}_Heatmap.png"))
    plt.close()

    # 2. HORIZONTAL BAR CHART (Fixed Logical Order)
    plt.figure(figsize=(10, 8))
    # Invert for horizontal display (top item in table = top bar)
    plot_data = overall_df.iloc[::-1]
    
    # Custom color logic
    colors = ['#1F4E79' if 'DistilBERT' in a else '#5B9BD5' for a in plot_data['Architecture']]
    
    bars = plt.barh(plot_data["Model Identifier"], plot_data["Accuracy"], color=colors)
    plt.xlabel("Overall Accuracy Score", fontsize=12)
    plt.title(f"Model Accuracy Comparison: {cond_name.replace('_', ' ')}", fontsize=15, pad=20)
    plt.xlim(0, 1.15) # Buffer for labels
    
    # Value labels
    for bar in bars:
        w = bar.get_width()
        plt.text(w + 0.01, bar.get_y() + bar.get_height()/2, f'{w:.4f}', va='center', fontweight='bold')
        
    plt.tight_layout()
    plt.savefig(os.path.join(out_dir, f"{cond_name}_Accuracy_Bars.png"))
    plt.close()

# =============================================================================
# EXCEL STYLING ENGINE
# =============================================================================

def apply_excel_styles(writer, sheet_name, df):
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    ws = writer.sheets[sheet_name]
    
    thin_side = Side(style='thin', color=_BORDER)
    full_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    
    # Header Styling
    h_fill = PatternFill(start_color=_HDR_BG, fill_type="solid")
    h_font = Font(color=_HDR_FG, bold=True)
    
    for col in range(1, len(df.columns) + 1):
        cell = ws.cell(row=1, column=col)
        cell.fill, cell.font, cell.border = h_fill, h_font, full_border
        cell.alignment = Alignment(horizontal="center")

    # Row Styling
    for r_idx, row in enumerate(ws.iter_rows(min_row=2), 2):
        bg_color = _ALT_ROW if r_idx % 2 == 0 else "FFFFFF"
        row_fill = PatternFill(start_color=bg_color, fill_type="solid")
        
        # Check for significance highlighting
        is_sig = False
        if "Significant" in df.columns:
            sig_col_idx = list(df.columns).index("Significant")
            if str(row[sig_col_idx].value).upper() == "TRUE":
                is_sig = True
                row_fill = PatternFill(start_color=_SIG_BG, fill_type="solid")

        for cell in row:
            cell.fill, cell.border = row_fill, full_border
            cell.alignment = Alignment(horizontal="center")

# =============================================================================
# MAIN PIPELINE
# =============================================================================

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--input_dir", required=True, help="Directory containing master_results.xlsx files")
    parser.add_argument("--output_dir", required=True, help="Directory for final tables and plots")
    args = parser.parse_args()

    os.makedirs(args.output_dir, exist_ok=True)
    final_excel = os.path.join(args.output_dir, "Thesis_Results_Full_Report.xlsx")

    # Pre-flight check for file lock
    if os.path.exists(final_excel):
        try:
            open(final_excel, 'a').close()
        except OSError:
            print(f"❌ ERROR: {final_excel} is open in Excel. Please close it and rerun.")
            return

    # 1. Load and Combine Data
    print("🚀 Loading and standardizing prediction data...")
    all_preds = []
    for folder in os.listdir(args.input_dir):
        path = os.path.join(args.input_dir, folder, "master_results.xlsx")
        if os.path.exists(path):
            preds = pd.read_excel(path, sheet_name="Predictions", dtype=str)
            all_preds.append(preds)
    
    if not all_preds:
        print("❌ No valid master_results.xlsx files found."); return
        
    full_df = pd.concat(all_preds, ignore_index=True)
    full_df['true_int'] = full_df['true_label'].map(LABEL_MAP)
    full_df['pred_int'] = full_df['pred_label'].map(LABEL_MAP)

    processed_sheets = {}

    # 2. Process by Thesis Condition
    for cond_id, hr_filter in CONDITIONS.items():
        print(f"📦 Processing {cond_id} ({hr_filter})...")
        
        # Filter data for this condition
        c_data = full_df[full_df['model_key'].str.contains(hr_filter, case=False)]
        if c_data.empty: continue
        
        overall_rows, subclass_rows = [], []
        
        # Metrics Calculation
        for (m_key, arch), group in c_data.groupby(["model_key", "arch"]):
            y_t, y_p = group['true_int'].values, group['pred_int'].values
            
            # Overall Table Metrics
            p, r, f1, _ = precision_recall_fscore_support(y_t, y_p, average='binary', zero_division=0)
            overall_rows.append({
                "Model Identifier": m_key, "Architecture": arch,
                "Accuracy": round(accuracy_score(y_t, y_p), 4),
                "Precision": round(p, 4), "Recall": round(r, 4), "F1-Score": round(f1, 4)
            })
            
            # Subclass Accuracy Metrics
            sc_meta = {"Model Identifier": m_key, "Architecture": arch}
            for sc in SUBCLASSES:
                sc_subset = group[group['subclass'] == sc]
                acc = accuracy_score(sc_subset['true_int'], sc_subset['pred_int']) if not sc_subset.empty else 0.0
                sc_meta[f"{sc} Accuracy"] = round(acc, 4)
            subclass_rows.append(sc_meta)

        # Create and Sort Final DataFrames
        o_df = standardize_and_sort(pd.DataFrame(overall_rows))
        s_df = standardize_and_sort(pd.DataFrame(subclass_rows))
        st_df = calculate_mcnemar_stats(c_data)

        processed_sheets[f"{cond_id}_Overall"] = o_df
        processed_sheets[f"{cond_id}_Subclass"] = s_df
        processed_sheets[f"{cond_id}_Stats"] = st_df
        
        # 3. Visualization
        print(f"  🎨 Generating charts for {cond_id}...")
        create_plots(o_df, s_df, cond_id, args.output_dir)

    # 4. Final Excel Generation
    print("📊 Writing stylized report to Excel...")
    with pd.ExcelWriter(final_excel, engine='openpyxl') as writer:
        for sheet, data in processed_sheets.items():
            data.to_excel(writer, sheet_name=sheet, index=False)
            apply_excel_styles(writer, sheet, data)

    print(f"\n✅ SUCCESS! All thesis outputs are ready in: {args.output_dir}")

if __name__ == "__main__":
    main()