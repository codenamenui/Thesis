"""
stratify.py
-----------
Produces a single Excel workbook where each sheet corresponds to one
news-type-ratio configuration, plus a Summary sheet.

Input — two modes (mutually exclusive):
  A) input_file   : path to a single Excel workbook whose sheets are already
                    merged (i.e. the output of merge_sheets.py).
  B) merge_sources: list of {file, sheet, rename?} entries — the same format
                    as merge_sheets.py's config — so you can skip the merge
                    step entirely and run just this script.

Split logic (applied to every sheet):
  - test  : 15% of total_samples, always equal across news types — FIXED
  - val   : 15% of total_samples, always equal across news types — FIXED
  - train : 70% of total_samples, at the sheet's configured news-type ratio

Test and val rows are sampled once and reused identically in every sheet.
Training rows are sampled fresh per configuration from the remaining pool.

Usage:
    python stratify.py config.json

Config keys:
    merge_sources   (option A) list of {file, sheet, rename?} — skips merge step
    input_file      (option B) path to already-merged Excel file
    sheet_names     (option B only, optional) map of news_type → sheet name
                    default: {"HR":"HR","HF":"HF","AI-F":"AI-F","AI-R":"AI-R"}
    output_file     (optional) default: stratified_dataset.xlsx
    seed            (optional) default: 42
    total_samples   (required)
    topic_ratios    (required) must sum to 1.0
    news_type_ratios (required) list of ratio dicts, each must sum to 1.0
"""

import json
import sys
import math
import pathlib
import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment


# ── config loading ────────────────────────────────────────────────────────────

def load_config(path: str) -> dict:
    with open(path) as f:
        cfg = json.load(f)

    if not cfg.get("merge_sources") and not cfg.get("input_file"):
        raise ValueError(
            "Config must have either 'merge_sources' (list) or 'input_file' (path)."
        )
    if not cfg.get("total_samples"):
        raise ValueError("'total_samples' is required in the config.")

    total = sum(cfg.get("topic_ratios", {}).values())
    if not math.isclose(total, 1.0, abs_tol=1e-6):
        raise ValueError(
            f"topic_ratios must sum to 1.0 (currently {total:.6f})."
        )
    return cfg


# ── data loading ──────────────────────────────────────────────────────────────

def _prep_df(df: pd.DataFrame, news_type: str, source_label: str) -> pd.DataFrame:
    """Normalise columns, validate, add news_type and _idx."""
    df = df.copy()
    df.columns = df.columns.str.strip().str.lower()

    missing = {"label", "article", "topic"} - set(df.columns)
    if missing:
        raise ValueError(f"'{source_label}' missing columns: {missing}")

    df["topic"]     = df["topic"].str.strip().str.lower()
    df["news_type"] = news_type
    df["_idx"]      = range(len(df))
    return df.reset_index(drop=True)


def load_from_sources(sources: list[dict]) -> dict[str, pd.DataFrame]:
    """
    Mode A — read directly from merge_sources entries.
    Each entry: {file, sheet, rename?}
    The news_type key is 'rename' if present, otherwise 'sheet'.
    """
    frames = {}
    for src in sources:
        file_path  = pathlib.Path(src["file"])
        sheet_name = src["sheet"]
        news_type  = src.get("rename", sheet_name)

        if not file_path.exists():
            raise FileNotFoundError(f"File not found: {file_path.resolve()}")

        xl  = pd.read_excel(file_path, sheet_name=sheet_name)
        df  = _prep_df(xl, news_type, f"{file_path.name}[{sheet_name}]")
        frames[news_type] = df
        print(f"  ✓  {file_path.name}[{sheet_name}]  →  '{news_type}'  ({len(df)} rows)")

    return frames


def load_from_file(input_file: str, sheet_map: dict) -> dict[str, pd.DataFrame]:
    """
    Mode B — read from a single merged Excel workbook.
    sheet_map: {news_type: sheet_name}
    """
    xl = pd.read_excel(input_file, sheet_name=None)
    frames = {}
    for news_type, sheet_name in sheet_map.items():
        if sheet_name not in xl:
            raise KeyError(
                f"Sheet '{sheet_name}' not found in '{input_file}'. "
                f"Available: {list(xl.keys())}"
            )
        df = _prep_df(xl[sheet_name], news_type, f"{input_file}[{sheet_name}]")
        frames[news_type] = df
        print(f"  ✓  {input_file}[{sheet_name}]  →  '{news_type}'  ({len(df)} rows)")

    return frames


# ── ratio → counts ────────────────────────────────────────────────────────────

def ratios_to_counts(ratios: dict[str, float], total: int) -> dict[str, int]:
    """Largest-remainder allocation so counts sum exactly to total."""
    raw    = {k: v * total for k, v in ratios.items()}
    floors = {k: int(v) for k, v in raw.items()}
    deficit = total - sum(floors.values())
    for k in sorted(raw, key=lambda k: raw[k] - floors[k], reverse=True)[:deficit]:
        floors[k] += 1
    return floors


def equal_ratio(news_types: list[str]) -> dict[str, float]:
    """Build an equal-weight ratio dict for any set of news types."""
    w = 1.0 / len(news_types)
    return {nt: w for nt in news_types}


# ── topic-stratified sampler ──────────────────────────────────────────────────

def sample_stratified(
    df: pd.DataFrame,
    n_total: int,
    topic_ratios: dict[str, float],
    rng: np.random.Generator,
    exclude_idx: set | None = None,
) -> pd.DataFrame:
    pool = df if not exclude_idx else df[~df["_idx"].isin(exclude_idx)]
    topic_counts = ratios_to_counts(topic_ratios, n_total)

    parts = []
    for topic, count in topic_counts.items():
        if count == 0:
            continue
        topic_pool = pool[pool["topic"] == topic]
        if topic_pool.empty:
            print(f"  [WARN] topic '{topic}' not found — skipped.")
            continue
        n_draw = min(count, len(topic_pool))
        if n_draw < count:
            print(f"  [WARN] topic '{topic}': wanted {count}, only {n_draw} available.")
        parts.append(
            topic_pool.sample(n=n_draw, random_state=int(rng.integers(1 << 31)))
        )

    return pd.concat(parts, ignore_index=True) if parts else pd.DataFrame(columns=df.columns)


# ── sheet name builder ────────────────────────────────────────────────────────

def sheet_label(type_ratios: dict[str, float]) -> str:
    """
    Shows each type's share within its group (real or fake), so all four
    percentages sum to 200.

    Real news group  (HR + AI-R): HR% + AI-R% = 100
    Fake news group  (HF + AI-F): HF% + AI-F% = 100
    Grand total                                = 200

    e.g. HR=0.165, AI-R=0.335, HF=0.335, AI-F=0.165
         real_total=0.50 → HR=33%, AI-R=67%
         fake_total=0.50 → HF=67%, AI-F=33%
         → 'HR33-AIR67-HF67-AIF33'
    """
    hr   = type_ratios.get("HR",   0)
    air  = type_ratios.get("AI-R", 0)
    hf   = type_ratios.get("HF",   0)
    aif  = type_ratios.get("AI-F", 0)

    real_total = hr  + air  or 1   # guard against div/0
    fake_total = hf  + aif  or 1

    hr_pct  = int(round(hr  / real_total * 100))
    air_pct = int(round(air / real_total * 100))
    hf_pct  = int(round(hf  / fake_total * 100))
    aif_pct = int(round(aif / fake_total * 100))

    return f"HR{hr_pct}-AIR{air_pct}-HF{hf_pct}-AIF{aif_pct}"

# ── Excel writing ─────────────────────────────────────────────────────────────

COL_ORDER   = ["label", "article", "topic", "news_type", "split"]
COL_WIDTHS  = {"label": 8, "article": 80, "topic": 20, "news_type": 12, "split": 10}

HEADER_FILL  = PatternFill("solid", start_color="1F4E79")
HEADER_FONT  = Font(name="Arial", bold=True, color="FFFFFF", size=11)
DATA_FONT    = Font(name="Arial", size=10)
SECTION_FONT = Font(name="Arial", bold=True, size=11)

SPLIT_COLORS = {
    "train": "E2EFDA",
    "val":   "FFF2CC",
    "test":  "FCE4D6",
}

# Summary sheet palette
SUMMARY_HEADER_FILL  = PatternFill("solid", start_color="1F4E79")
SUMMARY_SECTION_FILL = PatternFill("solid", start_color="D6E4F0")
SUMMARY_ALT_FILL     = PatternFill("solid", start_color="F2F7FB")
SUMMARY_TOTAL_FILL   = PatternFill("solid", start_color="BDD7EE")

NEWS_TYPE_COLORS = {
    "HR":   "C6EFCE",   # soft green
    "AI-R": "FFEB9C",   # soft yellow
    "HF":   "FFC7CE",   # soft red
    "AI-F": "E2AFFF",   # soft purple
}


def write_sheet(ws, df: pd.DataFrame) -> None:
    cols = [c for c in COL_ORDER if c in df.columns]
    df   = df[cols]

    for ci, col in enumerate(cols, 1):
        cell = ws.cell(row=1, column=ci, value=col.upper())
        cell.font      = HEADER_FONT
        cell.fill      = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for ri, row in enumerate(df.itertuples(index=False), 2):
        split_val = getattr(row, "split", "train")
        row_fill  = PatternFill("solid", start_color=SPLIT_COLORS.get(split_val, "FFFFFF"))
        for ci, val in enumerate(row, 1):
            cell = ws.cell(row=ri, column=ci, value=val)
            cell.font      = DATA_FONT
            cell.fill      = row_fill
            cell.alignment = Alignment(
                wrap_text=(cols[ci - 1] == "article"), vertical="top"
            )

    for ci, col in enumerate(cols, 1):
        ws.column_dimensions[ws.cell(row=1, column=ci).column_letter].width = \
            COL_WIDTHS.get(col, 15)

    ws.row_dimensions[1].height = 22
    ws.freeze_panes = "A2"


def _summary_header_row(ws, row: int, values: list, fill=None) -> None:
    """Write a styled header row to the summary sheet."""
    fill = fill or SUMMARY_HEADER_FILL
    for ci, val in enumerate(values, 1):
        cell = ws.cell(row=row, column=ci, value=val)
        cell.font      = HEADER_FONT
        cell.fill      = fill
        cell.alignment = Alignment(horizontal="center", vertical="center")


def _summary_cell(ws, row: int, col: int, value, bold=False,
                  fill=None, align="center") -> None:
    cell = ws.cell(row=row, column=col, value=value)
    cell.font      = Font(name="Arial", bold=bold, size=10)
    cell.alignment = Alignment(horizontal=align, vertical="center")
    if fill:
        cell.fill = fill


def write_summary_sheet(ws, summary_records: list[dict], news_types: list[str],
                        topic_ratios: dict, total_samples: int,
                        n_train: int, n_val: int, n_test: int) -> None:
    """
    Writes a multi-section summary sheet:

      §1  Global configuration  (seed, total, split targets, topic ratios)
      §2  Per-sheet overview    (one row per data sheet: train/val/test totals)
      §3  News-type breakdown   (per sheet × per split × per news type)
      §4  Topic breakdown       (per sheet × per split × per topic)
    """
    splits = ["train", "val", "test"]
    topics = sorted(topic_ratios.keys())
    ws.title = "Summary"

    # ── column widths ─────────────────────────────────────────────────────────
    ws.column_dimensions["A"].width = 28   # sheet name / label
    ws.column_dimensions["B"].width = 16   # sub-label
    for ci in range(3, 3 + max(len(news_types), len(topics)) + 2):
        ws.column_dimensions[
            ws.cell(row=1, column=ci).column_letter
        ].width = 13

    current_row = 1

    # ═════════════════════════════════════════════════════════════════════════
    # §1  GLOBAL CONFIGURATION
    # ═════════════════════════════════════════════════════════════════════════
    _summary_header_row(ws, current_row,
                        ["GLOBAL CONFIGURATION", "", "", ""],
                        fill=SUMMARY_HEADER_FILL)
    ws.merge_cells(start_row=current_row, start_column=1,
                   end_row=current_row, end_column=4)
    current_row += 1

    config_rows = [
        ("Total samples",  total_samples),
        ("Train (target)", n_train),
        ("Val   (target)", n_val),
        ("Test  (target)", n_test),
        ("Data sheets",    len(summary_records)),
    ]
    for label, val in config_rows:
        _summary_cell(ws, current_row, 1, label, bold=True, align="left",
                      fill=SUMMARY_ALT_FILL)
        _summary_cell(ws, current_row, 2, val, align="left",
                      fill=SUMMARY_ALT_FILL)
        current_row += 1

    # topic ratios sub-table
    _summary_cell(ws, current_row, 1, "Topic ratios", bold=True,
                  align="left", fill=SUMMARY_ALT_FILL)
    current_row += 1
    for topic, ratio in sorted(topic_ratios.items()):
        _summary_cell(ws, current_row, 1, f"  {topic}", align="left",
                      fill=SUMMARY_ALT_FILL)
        _summary_cell(ws, current_row, 2, f"{ratio:.1%}", align="left",
                      fill=SUMMARY_ALT_FILL)
        current_row += 1

    current_row += 1   # blank separator

    # ═════════════════════════════════════════════════════════════════════════
    # §2  PER-SHEET OVERVIEW
    # ═════════════════════════════════════════════════════════════════════════
    overview_cols = ["Sheet", "Train", "Val", "Test", "Total"]
    _summary_header_row(ws, current_row, overview_cols)
    # extend merged header label
    title_cell = ws.cell(row=current_row, column=1)
    title_cell.value = "PER-SHEET OVERVIEW"
    ws.merge_cells(start_row=current_row, start_column=1,
                   end_row=current_row, end_column=len(overview_cols))
    current_row += 1
    _summary_header_row(ws, current_row, overview_cols,
                        fill=SUMMARY_SECTION_FILL)
    for ci, h in enumerate(overview_cols, 1):
        ws.cell(row=current_row, column=ci).font = Font(
            name="Arial", bold=True, size=10)
    current_row += 1

    for rec in summary_records:
        alt = SUMMARY_ALT_FILL if summary_records.index(rec) % 2 else None
        _summary_cell(ws, current_row, 1, rec["sheet"], bold=True,
                      align="left", fill=alt)
        _summary_cell(ws, current_row, 2, rec["n_train"], fill=alt)
        _summary_cell(ws, current_row, 3, rec["n_val"],   fill=alt)
        _summary_cell(ws, current_row, 4, rec["n_test"],  fill=alt)
        _summary_cell(ws, current_row, 5,
                      rec["n_train"] + rec["n_val"] + rec["n_test"],
                      bold=True, fill=SUMMARY_TOTAL_FILL)
        current_row += 1

    current_row += 1

    # ═════════════════════════════════════════════════════════════════════════
    # §3  NEWS-TYPE BREAKDOWN  (sheet × split × news_type)
    # ═════════════════════════════════════════════════════════════════════════
    nt_cols = ["Sheet", "Split"] + news_types + ["Row Total"]
    title_cell = ws.cell(row=current_row, column=1)
    title_cell.value = "NEWS-TYPE BREAKDOWN"
    title_cell.font  = HEADER_FONT
    title_cell.fill  = SUMMARY_HEADER_FILL
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.merge_cells(start_row=current_row, start_column=1,
                   end_row=current_row, end_column=len(nt_cols))
    current_row += 1

    _summary_header_row(ws, current_row, nt_cols, fill=SUMMARY_SECTION_FILL)
    for ci, h in enumerate(nt_cols, 1):
        ws.cell(row=current_row, column=ci).font = Font(
            name="Arial", bold=True, size=10)
    current_row += 1

    for rec in summary_records:
        first_split = True
        for split in splits:
            nt_counts = rec["news_type_counts"].get(split, {})
            row_total = sum(nt_counts.values())
            alt = SUMMARY_ALT_FILL if summary_records.index(rec) % 2 else None

            # sheet name only on first split row, then blank
            _summary_cell(ws, current_row, 1,
                          rec["sheet"] if first_split else "",
                          bold=first_split, align="left", fill=alt)
            split_fill = PatternFill("solid",
                                     start_color=SPLIT_COLORS.get(split, "FFFFFF"))
            _summary_cell(ws, current_row, 2, split, fill=split_fill)

            for ci, nt in enumerate(news_types, 3):
                nt_fill = PatternFill("solid",
                                      start_color=NEWS_TYPE_COLORS.get(nt, "FFFFFF"))
                _summary_cell(ws, current_row, ci,
                               nt_counts.get(nt, 0), fill=nt_fill)
            _summary_cell(ws, current_row, 3 + len(news_types),
                          row_total, bold=True, fill=SUMMARY_TOTAL_FILL)
            current_row += 1
            first_split = False

    current_row += 1

    # ═════════════════════════════════════════════════════════════════════════
    # §4  TOPIC BREAKDOWN  (sheet × split × topic)
    # ═════════════════════════════════════════════════════════════════════════
    topic_cols = ["Sheet", "Split"] + topics + ["Row Total"]
    title_cell = ws.cell(row=current_row, column=1)
    title_cell.value = "TOPIC BREAKDOWN"
    title_cell.font  = HEADER_FONT
    title_cell.fill  = SUMMARY_HEADER_FILL
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.merge_cells(start_row=current_row, start_column=1,
                   end_row=current_row, end_column=len(topic_cols))
    current_row += 1

    _summary_header_row(ws, current_row, topic_cols, fill=SUMMARY_SECTION_FILL)
    for ci, h in enumerate(topic_cols, 1):
        ws.cell(row=current_row, column=ci).font = Font(
            name="Arial", bold=True, size=10)
    current_row += 1

    for rec in summary_records:
        first_split = True
        for split in splits:
            topic_counts = rec["topic_counts"].get(split, {})
            row_total    = sum(topic_counts.values())
            alt = SUMMARY_ALT_FILL if summary_records.index(rec) % 2 else None

            _summary_cell(ws, current_row, 1,
                          rec["sheet"] if first_split else "",
                          bold=first_split, align="left", fill=alt)
            split_fill = PatternFill("solid",
                                     start_color=SPLIT_COLORS.get(split, "FFFFFF"))
            _summary_cell(ws, current_row, 2, split, fill=split_fill)

            for ci, topic in enumerate(topics, 3):
                _summary_cell(ws, current_row, ci,
                               topic_counts.get(topic, 0), fill=alt)
            _summary_cell(ws, current_row, 3 + len(topics),
                          row_total, bold=True, fill=SUMMARY_TOTAL_FILL)
            current_row += 1
            first_split = False

    ws.freeze_panes = "A2"


# ── main ──────────────────────────────────────────────────────────────────────

def main(config_path: str) -> None:
    cfg           = load_config(config_path)
    seed          = cfg.get("seed", 42)
    output_file   = cfg.get("output_file", "stratified_dataset.xlsx")
    total_samples = cfg["total_samples"]
    topic_ratios  = cfg["topic_ratios"]
    ratio_configs = cfg["news_type_ratios"]

    split_counts  = ratios_to_counts({"test": 0.15, "val": 0.15, "train": 0.70}, total_samples)
    n_tv  = split_counts["test"]
    n_vv  = split_counts["val"]
    n_tr  = split_counts["train"]

    print(f"\n📄  Output        : {output_file}")
    print(f"🌱  Seed          : {seed}")
    print(f"📦  Total samples : {total_samples}  "
          f"(train={n_tr}, val={n_vv}, test={n_tv})")
    print(f"📊  Topic ratios  : {topic_ratios}")
    print(f"🔢  Ratio configs : {len(ratio_configs)}\n")

    # ── load data ─────────────────────────────────────────────────────────────
    if "merge_sources" in cfg:
        print("Loading from merge_sources …")
        frames = load_from_sources(cfg["merge_sources"])
    else:
        print(f"Loading from input_file: {cfg['input_file']} …")
        sheet_map = cfg.get(
            "sheet_names",
            {"HR": "HR", "HF": "HF", "AI-F": "AI-F", "AI-R": "AI-R"}
        )
        frames = load_from_file(cfg["input_file"], sheet_map)

    news_types   = sorted(frames.keys())
    eq_ratio     = equal_ratio(news_types)
    print(f"\n  News types found : {news_types}")
    print(f"  Equal test/val ratio : {eq_ratio}\n")

    # ── sample fixed test + val (equal across news types, topic-stratified) ───
    print("Sampling fixed test / val sets …")
    rng_fixed = np.random.default_rng(seed)
    tv_counts = ratios_to_counts(eq_ratio, n_tv)
    vv_counts = ratios_to_counts(eq_ratio, n_vv)

    test_parts, val_parts = [], []
    used: dict[str, set] = {nt: set() for nt in frames}

    for news_type, count in tv_counts.items():
        part = sample_stratified(
            frames[news_type], count, topic_ratios, rng_fixed,
            exclude_idx=used[news_type]
        )
        part["split"] = "test"
        used[news_type].update(part["_idx"].tolist())
        test_parts.append(part)

    for news_type, count in vv_counts.items():
        part = sample_stratified(
            frames[news_type], count, topic_ratios, rng_fixed,
            exclude_idx=used[news_type]
        )
        part["split"] = "val"
        used[news_type].update(part["_idx"].tolist())
        val_parts.append(part)

    test_df = pd.concat(test_parts, ignore_index=True)
    val_df  = pd.concat(val_parts,  ignore_index=True)
    print(f"  test rows : {len(test_df)}")
    print(f"  val rows  : {len(val_df)}\n")

    # ── build workbook ────────────────────────────────────────────────────────
    wb = Workbook()
    wb.remove(wb.active)

    summary_records: list[dict] = []

    for idx, type_ratios in enumerate(ratio_configs, 1):
        ratio_sum = sum(type_ratios.values())
        if not math.isclose(ratio_sum, 1.0, abs_tol=1e-6):
            raise ValueError(
                f"news_type_ratios[{idx-1}] must sum to 1.0 "
                f"(currently {ratio_sum:.6f})."
            )

        name = sheet_label(type_ratios)
        print(f"[{idx}/{len(ratio_configs)}] Sheet '{name}'  train ratio={type_ratios}")

        train_counts = ratios_to_counts(type_ratios, n_tr)
        train_parts  = []
        for news_type, count in train_counts.items():
            rng_train = np.random.default_rng(seed + idx * 1000)
            part = sample_stratified(
                frames[news_type], count, topic_ratios, rng_train,
                exclude_idx=used[news_type]
            )
            part["split"] = "train"
            train_parts.append(part)

        train_df = pd.concat(train_parts, ignore_index=True)
        train_df = train_df.sample(frac=1, random_state=seed).reset_index(drop=True)
        combined = pd.concat([train_df, val_df, test_df], ignore_index=True)
        combined.drop(columns=["_idx"], inplace=True)

        print(f"  rows — train:{len(train_df)}  val:{len(val_df)}  test:{len(test_df)}  "
              f"total:{len(combined)}")
        print(f"  news-type dist (train):\n"
              f"{train_df['news_type'].value_counts().to_string()}\n")

        # ── collect stats for Summary sheet ───────────────────────────────────
        nt_counts: dict[str, dict[str, int]] = {}
        tp_counts: dict[str, dict[str, int]] = {}
        for split_name, split_df in [("train", train_df),
                                      ("val",   val_df),
                                      ("test",  test_df)]:
            nt_counts[split_name] = (
                split_df["news_type"].value_counts().to_dict()
            )
            tp_counts[split_name] = (
                split_df["topic"].value_counts().to_dict()
            )

        summary_records.append({
            "sheet":            name,
            "n_train":          len(train_df),
            "n_val":            len(val_df),
            "n_test":           len(test_df),
            "news_type_counts": nt_counts,
            "topic_counts":     tp_counts,
        })

        ws = wb.create_sheet(title=name)
        write_sheet(ws, combined)

    # ── append Summary sheet ──────────────────────────────────────────────────
    print("Writing Summary sheet …")
    ws_summary = wb.create_sheet(title="Summary")
    write_summary_sheet(
        ws_summary,
        summary_records,
        news_types=news_types,
        topic_ratios=topic_ratios,
        total_samples=total_samples,
        n_train=n_tr,
        n_val=n_vv,
        n_test=n_tv,
    )

    out_path = pathlib.Path(output_file)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(out_path))
    print(f"✅  Saved → {out_path}")


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python stratify.py config.json")
        sys.exit(1)
    main(sys.argv[1])