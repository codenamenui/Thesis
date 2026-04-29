#!/usr/bin/env python3
"""
merge_results_thesis.py
==============================================
Generates structured tables specifically for Thesis Chapters:
Condition A (HR100), Condition B (HR67), and Condition C (HR50).
"""

import argparse
import os
import re
import warnings
from typing import Dict, List, Tuple

import numpy as np
import pandas as pd
from sklearn.metrics import precision_recall_fscore_support, accuracy_score

warnings.filterwarnings("ignore")

# =============================================================================
# CONSTANTS & CONDITION MAPPING
# =============================================================================

SUBCLASSES = ["HR", "AI-R", "HF", "AI-F"]
LABEL_STR_TO_INT = {"real": 0, "fake": 1}

# Condition Mapping based on HR Ratio
CONDITIONS = {
    "Condition_A": "HR100",
    "Condition_B": "HR67",
    "Condition_C": "HR50"
}

# Styling
_HDR_BG, _HDR_FG = "1F4E79", "FFFFFF"
_ALT_ROW = "D6E4F0"
_SIG = "FFEB9C"
_BORDER = "BDD7EE"

# =============================================================================
# HELPERS
# =============================================================================

def _get_sort_key(key: str) -> Tuple:
    key_str = str(key).lower()
    arch_val = 1 if "distilbert" in key_str else 0
    hf_match = re.search(r'hf(\d+)', key_str)
    hf_val = int(hf_match.group(1)) if hf_match else 0
    return (arch_val, -hf_val)

def _standardize_df(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty: return df
    if 'arch' in df.columns:
        df['arch'] = df['arch'].replace({'BERT': 'Tagalog-BERT', 'DISTILBERT': 'Tagalog-DistilBERT'})
    if 'model_key' in df.columns:
        df['_sort'] = df['model_key'].apply(_get_sort_key)
        df = df.sort_values('_sort').drop(columns=['_sort'])
    return df

# =============================================================================
# TABLE GENERATOR
# =============================================================================

def generate_thesis_tables(predictions_df: pd.DataFrame) -> Dict[str, pd.DataFrame]:
    if predictions_df.empty: return {}
    
    df = predictions_df.copy()
    df['true_int'] = df['true_label'].map(LABEL_STR_TO_INT)
    df['pred_int'] = df['pred_label'].map(LABEL_STR_TO_INT)
    
    tables = {}

    for cond_name, hr_tag in CONDITIONS.items():
        # Filter data for this specific condition
        cond_data = df[df['model_key'].str.contains(hr_tag, case=False)]
        if cond_data.empty: continue
        
        overall_rows = []
        subclass_rows = []
        
        grouped = cond_data.groupby(["model_key", "arch"])
        for (m_key, arch), group in grouped:
            # Table X: Overall Performance
            y_t, y_p = group['true_int'].values, group['pred_int'].values
            p, r, f1, _ = precision_recall_fscore_support(y_t, y_p, average='binary', zero_division=0)
            overall_rows.append({
                "Model Identifier": m_key,
                "Architecture": arch,
                "Accuracy": round(accuracy_score(y_t, y_p), 4),
                "Precision": round(p, 4),
                "Recall": round(r, 4),
                "F1-Score": round(f1, 4)
            })
            
            # Table X: Subclass-Wise Accuracy
            sc_row = {"Model Identifier": m_key, "Architecture": arch}
            for sc in SUBCLASSES:
                sc_g = group[group['subclass'] == sc]
                acc = accuracy_score(sc_g['true_int'], sc_g['pred_int']) if not sc_g.empty else 0.0
                sc_row[f"{sc} Accuracy"] = round(acc, 4)
            subclass_rows.append(sc_row)
            
        tables[f"{cond_name}_Overall"] = _standardize_df(pd.DataFrame(overall_rows))
        tables[f"{cond_name}_Subclass"] = _standardize_df(pd.DataFrame(subclass_rows))

    return tables

# =============================================================================
# STATS ENGINE (Thesis Specific)
# =============================================================================

def run_thesis_stats(predictions_df: pd.DataFrame) -> Dict[str, pd.DataFrame]:
    from statsmodels.stats.multitest import multipletests
    from statsmodels.stats.contingency_tables import mcnemar
    
    stats_tables = {}
    for cond_name, hr_tag in CONDITIONS.items():
        cond_data = predictions_df[predictions_df['model_key'].str.contains(hr_tag, case=False)]
        models = cond_data['model_key'].unique()
        res = []
        for i in range(len(models)):
            for j in range(i + 1, len(models)):
                m1, m2 = models[i], models[j]
                d1 = cond_data[cond_data['model_key'] == m1].sort_values('sample_index')
                d2 = cond_data[cond_data['model_key'] == m2].sort_values('sample_index')
                common = min(len(d1), len(d2))
                y, p1, p2 = d1['true_label'].values[:common], d1['pred_label'].values[:common], d2['pred_label'].values[:common]
                b, c = np.sum((p1 == y) & (p2 != y)), np.sum((p1 != y) & (p2 == y))
                p_val = mcnemar([[0, b], [c, 0]], exact=(b + c < 25)).pvalue
                res.append({"Comparison": f"{m1} vs {m2}", "b": b, "c": c, "p-value": p_val})
        
        if res:
            res_df = pd.DataFrame(res)
            _, p_adj, _, _ = multipletests(res_df['p-value'], alpha=0.05, method='fdr_bh')
            res_df['p-adj (BH)'] = p_adj
            res_df['Significant'] = p_adj < 0.05
            stats_tables[f"{cond_name}_Stats"] = res_df
            
    return stats_tables

# =============================================================================
# EXCEL FORMATTING
# =============================================================================

def apply_thesis_styles(writer, sheet, df):
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    ws = writer.sheets[sheet]
    thin = Side(style='thin', color=_BORDER)
    brd = Border(left=thin, right=thin, top=thin, bottom=thin)
    
    for c in range(1, len(df.columns) + 1):
        cell = ws.cell(1, c)
        cell.fill, cell.font = PatternFill("solid", start_color=_HDR_BG), Font(color=_HDR_FG, bold=True)
        cell.alignment, cell.border = Alignment(horizontal="center"), brd

    for r_idx, row in enumerate(ws.iter_rows(min_row=2), 2):
        bg = _ALT_ROW if r_idx % 2 == 0 else "FFFFFF"
        for cell in row:
            cell.fill, cell.border, cell.alignment = PatternFill("solid", start_color=bg), brd, Alignment(horizontal="center")
        
        if "Stats" in sheet and str(row[list(df.columns).index("Significant")].value).lower() == "true":
            for cell in row: cell.fill = PatternFill("solid", start_color=_SIG)

# =============================================================================
# MAIN
# =============================================================================

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--input_dir", required=True)
    parser.add_argument("--output_dir", required=True)
    args = parser.parse_args()

    out_path = os.path.join(args.output_dir, "Thesis_Results_Tables.xlsx")
    
    # Permission check
    if os.path.exists(out_path):
        try: open(out_path, 'a').close()
        except OSError:
            print("❌ Close Excel first!"); return

    # Load All Data
    all_preds = []
    for f in os.listdir(args.input_dir):
        p = os.path.join(args.input_dir, f, "master_results.xlsx")
        if os.path.exists(p):
            df = pd.read_excel(p, sheet_name="Predictions", dtype=str)
            all_preds.append(df)
    
    if not all_preds: print("No data!"); return
    full_df = pd.concat(all_preds, ignore_index=True)

    # Process Tables
    perf_tables = generate_thesis_tables(full_df)
    stat_tables = run_thesis_stats(full_df)

    # Save
    os.makedirs(args.output_dir, exist_ok=True)
    with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
        for name, data in {**perf_tables, **stat_tables}.items():
            data.to_excel(writer, sheet_name=name, index=False)
            apply_thesis_styles(writer, name, data)

    print(f"\n[SUCCESS] Thesis tables generated at: {out_path}")

if __name__ == "__main__":
    main()