"""
length_pipeline.py
------------------
Full pipeline in one script:

    1. TOKENIZE  — tokenize raw text using jcblaise/bert-tagalog-base-cased
    2. TRUNCATE  — cap all articles to a single global token threshold
    3. SIMULATE  — resample AI-F → HF distribution, AI-R → HR distribution
    4. CLASSIFY  — logistic regression on token_count; report test accuracy only

Global threshold
----------------
Computed as the median token count of whichever of the four categories
(HR, HF, AI-R, AI-F) has the lowest overall median across ALL sheets and
ALL splits combined. This makes the threshold data-driven and reproducible.

Usage
-----
    pip install transformers torch openpyxl pandas numpy scikit-learn
    python length_pipeline.py stratified_dataset.xlsx
    # Output: length_pipeline_results.xlsx
"""

import sys
import pathlib
import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from sklearn.linear_model import LogisticRegression
from sklearn.metrics import accuracy_score

# ── Constants ─────────────────────────────────────────────────────────────────

TAGALOG_BERT_MODEL   = "jcblaise/bert-tagalog-base-cased"
TEXT_COL_CANDIDATES  = ["text", "content", "article", "body", "sentence", "news"]
LABEL_MAP            = {"HR": 1, "AI-R": 1, "HF": 0, "AI-F": 0}
SEED                 = 42
rng                  = np.random.default_rng(SEED)

HDR_FILL    = "1F4E79"
GREEN_FILL  = "C6EFCE"
BLUE_FILL   = "BDD7EE"
YELLOW_FILL = "FFEB9C"
RED_FILL    = "FFC7CE"
GREY_FILL   = "D9D9D9"
ALT_ROW     = "F2F2F2"
NT_FILLS    = {"HR": "C6EFCE", "AI-R": "FFEB9C", "HF": "FFC7CE", "AI-F": "E2AFFF"}

# ── Tokenizer ─────────────────────────────────────────────────────────────────

_tokenizer = None

def get_tokenizer():
    global _tokenizer
    if _tokenizer is None:
        from transformers import AutoTokenizer
        print(f"\n  [1/4] Loading tokenizer: {TAGALOG_BERT_MODEL} …")
        _tokenizer = AutoTokenizer.from_pretrained(TAGALOG_BERT_MODEL)
        print("        Tokenizer ready.")
    return _tokenizer


def tokenize_texts(texts: list[str], batch_size: int = 128) -> list[list[int]]:
    tok     = get_tokenizer()
    all_ids = []
    for i in range(0, len(texts), batch_size):
        batch = texts[i : i + batch_size]
        enc   = tok(batch, add_special_tokens=True, truncation=False,
                    padding=False, return_attention_mask=False)
        all_ids.extend(enc["input_ids"])
        if i > 0 and i % 2048 == 0:
            print(f"        Tokenized {i}/{len(texts)} rows …")
    return all_ids


def decode_ids(ids: list[int]) -> str:
    return get_tokenizer().decode(ids, skip_special_tokens=True)


def find_text_col(cols: list[str]) -> str | None:
    lower = [c.lower() for c in cols]
    candidates = TEXT_COL_CANDIDATES + ["article"]
    for cand in candidates:
        if cand in lower:
            return cols[lower.index(cand)]
    skip = {"label", "split", "news_type", "id", "index", "token_count",
            "original_article", "topic", "_ids", "_text_col"}
    others = [c for c in cols if c.lower() not in skip]
    return others[0] if others else None


# ── Step 1 — Tokenize ─────────────────────────────────────────────────────────

def step_tokenize(sheets: dict[str, pd.DataFrame]) -> dict[str, pd.DataFrame]:
    """Add token_count and _ids columns to every sheet."""
    print("\n  [1/4] TOKENIZE")
    out = {}
    for sheet_name, df in sheets.items():
        text_col = find_text_col(list(df.columns))
        if text_col is None:
            print(f"        [SKIP] '{sheet_name}' — no text column found.")
            continue
        print(f"        '{sheet_name}' — {len(df)} rows …")
        texts        = df[text_col].fillna("").astype(str).tolist()
        all_ids      = tokenize_texts(texts)
        df           = df.copy()
        df["_ids"]   = all_ids
        df["token_count"] = [len(ids) for ids in all_ids]
        df["_text_col"]   = text_col
        out[sheet_name]   = df
    return out


# ── Step 2 — Truncate ─────────────────────────────────────────────────────────

def compute_global_threshold(sheets: dict[str, pd.DataFrame]) -> tuple[int, str]:
    """
    Global cap = MEAN token count of the category with the lowest mean,
    computed from TRAIN rows only across all sheets.
    Mean is used instead of median — it is less aggressive, pulled up by
    longer articles, and preserves more natural variance in the distribution.
    Train-only prevents test data from influencing the threshold.
    """
    category_tokens: dict[str, list] = {}
    for df in sheets.values():
        train_df = df[df["split"] == "train"]
        for nt in ["HR", "HF", "AI-R", "AI-F"]:
            tokens = train_df[train_df["news_type"] == nt]["token_count"].values
            if len(tokens) == 0:
                continue
            category_tokens.setdefault(nt, [])
            category_tokens[nt].extend(tokens.tolist())

    if not category_tokens:
        raise ValueError("No train rows found — cannot compute threshold.")

    means       = {nt: float(np.mean(vals)) for nt, vals in category_tokens.items()}
    shortest_nt = min(means, key=means.get)
    threshold   = int(means[shortest_nt])
    description = (f"mean of {shortest_nt} (train only) across all sheets "
                   f"({', '.join(f'{k}={v:.0f}' for k, v in sorted(means.items()))})")
    return threshold, description


def truncate_ids(ids: list[int], cap: int) -> list[int]:
    if len(ids) <= cap:
        return ids
    return ids[:cap - 1] + [ids[-1]]   # preserve [CLS] … [SEP]


def step_truncate(sheets: dict[str, pd.DataFrame]) -> tuple[dict[str, pd.DataFrame], int, str]:
    """
    Truncate TRAIN articles to the global threshold.
    TEST articles are left at their natural length — they must not be
    modified, since the test set represents real-world article lengths
    and should not be artificially shortened.
    """
    cap, cap_desc = compute_global_threshold(sheets)
    print(f"\n  [2/4] TRUNCATE  →  global cap = {cap} tokens (train only)")
    print(f"        ({cap_desc})")

    out = {}
    for sheet_name, df in sheets.items():
        df       = df.copy()
        text_col = df["_text_col"].iloc[0]

        is_train = df["split"] == "train"

        # Truncate train rows
        train_trunc = df.loc[is_train, "_ids"].apply(lambda ids: truncate_ids(ids, cap))
        df.loc[is_train, text_col]      = train_trunc.apply(decode_ids)
        df.loc[is_train, "token_count"] = train_trunc.apply(len)
        df.loc[is_train, "_ids"]        = train_trunc

        pct = (df.loc[is_train, "token_count"] == cap).mean()
        n_test = (~is_train).sum()
        print(f"        '{sheet_name}': {pct:.1%} of train rows truncated  |  {n_test} test rows untouched")
        out[sheet_name] = df
    return out, cap, cap_desc


# ── Step 3 — Simulate ─────────────────────────────────────────────────────────

def _clamp(arr: np.ndarray, lo: int, hi: int) -> np.ndarray:
    return np.clip(np.round(arr).astype(int), max(lo, 1), hi)


def _sample_like(ref: np.ndarray, n: int) -> np.ndarray:
    if len(ref) == 0 or n == 0:
        return np.array([], dtype=int)
    mean, std = float(ref.mean()), float(ref.std())
    lo,   hi  = int(ref.min()), int(ref.max())
    if std < 1:
        return rng.choice(ref, size=n, replace=True)
    samples  = _clamp(rng.normal(mean, std, size=n * 4), lo, hi)
    in_range = samples[(samples >= lo) & (samples <= hi)]
    if len(in_range) < n:
        in_range = np.concatenate(
            [in_range, rng.choice(ref, size=n - len(in_range), replace=True)]
        )
    return in_range[:n]


def step_simulate(sheets: dict[str, pd.DataFrame],
                   natural_sheets: dict[str, pd.DataFrame]) -> dict[str, pd.DataFrame]:
    """
    For each sheet + split, resample:
      AI-F token counts → HF distribution
      AI-R token counts → HR distribution
    HR and HF are left unchanged.

    Reference distributions are taken from natural_sheets (pre-truncation),
    separately per split. This ensures the test simulation uses real test
    lengths as reference, not the truncated train lengths.

    Returns sim DataFrames with only: token_count, label, split, news_type.
    """
    print("\n  [3/4] SIMULATE")
    sim_sheets: dict[str, pd.DataFrame] = {}

    for sheet_name, df in sheets.items():
        nat_df = natural_sheets.get(sheet_name, df)
        rows = []
        for split, sdf in df.groupby("split"):
            nat_sdf   = nat_df[nat_df["split"] == split]
            hf_tokens = nat_sdf.loc[nat_sdf["news_type"] == "HF", "token_count"].values
            hr_tokens = nat_sdf.loc[nat_sdf["news_type"] == "HR", "token_count"].values

            for nt, grp in sdf.groupby("news_type"):
                n    = len(grp)
                orig = grp["token_count"].values

                if nt == "AI-F":
                    new = _sample_like(hf_tokens, n) if len(hf_tokens) else orig
                elif nt == "AI-R":
                    new = _sample_like(hr_tokens, n) if len(hr_tokens) else orig
                else:
                    new = orig

                label = LABEL_MAP.get(nt, 0)
                for tok in new:
                    rows.append({"token_count": int(tok), "label": label,
                                 "split": split, "news_type": nt})

        sim_sheets[sheet_name] = pd.DataFrame(
            rows, columns=["token_count", "label", "split", "news_type"]
        )
        print(f"        '{sheet_name}': {len(rows)} rows simulated")

    return sim_sheets


# ── Step 4 — Classify ─────────────────────────────────────────────────────────

def _run_classifier(df: pd.DataFrame, sheet_name: str) -> dict:
    train = df[df["split"] == "train"]
    test  = df[df["split"] == "test"]

    if len(train) == 0 or len(test) == 0:
        return None

    # Guard: if only one class in train, classifier is meaningless
    if train["label"].nunique() < 2:
        print(f"        [WARN] '{sheet_name}' train has only one class — skipping.")
        return None

    clf = LogisticRegression(random_state=SEED, max_iter=1000)
    clf.fit(train[["token_count"]], train["label"])

    return {
        "sheet":        sheet_name,
        "train_n":      len(train),
        "test_n":       len(test),
        "train_acc":    accuracy_score(train["label"], clf.predict(train[["token_count"]])),
        "test_acc":     accuracy_score(test["label"],  clf.predict(test[["token_count"]])),
        "coef":         float(clf.coef_[0][0]),
        "intercept":    float(clf.intercept_[0]),
    }


def step_classify(
    orig_sheets: dict[str, pd.DataFrame],
    sim_sheets:  dict[str, pd.DataFrame],
) -> tuple[list[dict], list[dict]]:
    print("\n  [4/4] CLASSIFY")
    before, after = [], []

    for sheet_name, df in orig_sheets.items():
        r = _run_classifier(df[["token_count", "label", "split"]], sheet_name)
        if r:
            before.append(r)
            print(f"        BEFORE '{sheet_name}': test acc = {r['test_acc']:.4f}")

    for sheet_name, sim_df in sim_sheets.items():
        r = _run_classifier(sim_df, sheet_name)
        if r:
            after.append(r)
            print(f"        AFTER  '{sheet_name}': test acc = {r['test_acc']:.4f}")

    return before, after


# ── Excel helpers ─────────────────────────────────────────────────────────────

def _hdr(ws, row, col, value):
    c = ws.cell(row=row, column=col, value=value)
    c.font      = Font(name="Arial", bold=True, size=10, color="FFFFFF")
    c.fill      = PatternFill("solid", start_color=HDR_FILL)
    c.alignment = Alignment(horizontal="center", vertical="center")


def _cell(ws, row, col, value, fill_hex=None, bold=False, fmt=None, align="center"):
    c = ws.cell(row=row, column=col, value=value)
    c.font      = Font(name="Arial", bold=bold, size=10)
    c.alignment = Alignment(horizontal=align, vertical="center")
    if fill_hex:
        c.fill = PatternFill("solid", start_color=fill_hex)
    if fmt:
        c.number_format = fmt


def _widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


# ── Excel writers ─────────────────────────────────────────────────────────────

def write_accuracy_sheet(ws, before: list[dict], after: list[dict],
                         cap: int, cap_desc: str) -> None:
    ws.title = "Test Accuracy"

    headers = ["Sheet", "Before Test Acc", "After Test Acc", "Δ Acc"]
    _widths(ws, [42, 18, 17, 12])
    for ci, h in enumerate(headers, 1):
        _hdr(ws, 1, ci, h)
    ws.row_dimensions[1].height = 28

    b_map = {r["sheet"]: r for r in before}
    a_map = {r["sheet"]: r for r in after}

    ri = 2
    for sheet_name in b_map:
        b = b_map[sheet_name]
        a = a_map.get(sheet_name)
        if a is None:
            continue
        delta     = a["test_acc"] - b["test_acc"]
        d_fill    = GREEN_FILL if delta < -0.05 else (RED_FILL if delta > 0.05 else GREY_FILL)
        row_fill  = ALT_ROW if ri % 2 == 0 else None
        _cell(ws, ri, 1, sheet_name,              bold=True, align="left")
        _cell(ws, ri, 2, b["test_acc"],            fill_hex=YELLOW_FILL, fmt="0.00%")
        _cell(ws, ri, 3, a["test_acc"],            fill_hex=BLUE_FILL,   fmt="0.00%")
        _cell(ws, ri, 4, delta,                    fill_hex=d_fill,
              fmt="+0.00%;-0.00%;0.00%")
        ri += 1

    # Summary row
    b_accs = [b_map[s]["test_acc"] for s in b_map if s in a_map]
    a_accs = [a_map[s]["test_acc"] for s in b_map if s in a_map]
    ri += 1
    _cell(ws, ri, 1, "MEAN",          bold=True)
    _cell(ws, ri, 2, np.mean(b_accs), bold=True, fill_hex=YELLOW_FILL, fmt="0.00%")
    _cell(ws, ri, 3, np.mean(a_accs), bold=True, fill_hex=BLUE_FILL,   fmt="0.00%")
    delta_mean = np.mean(a_accs) - np.mean(b_accs)
    _cell(ws, ri, 4, delta_mean, bold=True,
          fill_hex=GREEN_FILL if delta_mean < 0 else RED_FILL,
          fmt="+0.00%;-0.00%;0.00%")

    # Threshold note
    ri += 2
    note = ws.cell(row=ri, column=1,
                   value=f"Global truncation cap: {cap} tokens  |  {cap_desc}")
    note.font = Font(name="Arial", italic=True, size=9, color="595959")

    ri += 1
    interp = ws.cell(row=ri, column=1,
                     value="Δ Acc < 0 (green) = length became less predictive after truncation + simulation — desired outcome.")
    interp.font = Font(name="Arial", italic=True, size=9, color="375623")

    ws.freeze_panes = "A2"


def write_token_stats_sheet(ws, sheets: dict[str, pd.DataFrame], title: str) -> None:
    ws.title = title[:31]
    headers  = ["Sheet", "Split", "News Type", "n", "Mean", "Std", "Min", "Max", "Median"]
    _widths(ws, [42, 10, 12, 8, 10, 10, 8, 8, 10])
    for ci, h in enumerate(headers, 1):
        _hdr(ws, 1, ci, h)
    ws.row_dimensions[1].height = 28

    SPLIT_FILLS = {"train": GREEN_FILL, "test": BLUE_FILL}
    ri = 2
    for sheet_name, df in sheets.items():
        for split in ["train", "test"]:
            sdf   = df[df["split"] == split]
            first = True
            for nt in ["HR", "AI-R", "HF", "AI-F"]:
                grp = sdf[sdf["news_type"] == nt]["token_count"]
                if len(grp) == 0:
                    continue
                fill = NT_FILLS.get(nt, ALT_ROW)
                _cell(ws, ri, 1, sheet_name if first else "", bold=first, align="left")
                _cell(ws, ri, 2, split if first else "",
                      fill_hex=SPLIT_FILLS.get(split) if first else None)
                _cell(ws, ri, 3, nt,                   fill_hex=fill)
                _cell(ws, ri, 4, len(grp),             fill_hex=fill)
                _cell(ws, ri, 5, round(grp.mean(), 1), fill_hex=fill, fmt="0.0")
                _cell(ws, ri, 6, round(grp.std(),  1), fill_hex=fill, fmt="0.0")
                _cell(ws, ri, 7, int(grp.min()),       fill_hex=fill)
                _cell(ws, ri, 8, int(grp.max()),       fill_hex=fill)
                _cell(ws, ri, 9, round(grp.median(),1),fill_hex=fill, fmt="0.0")
                ri   += 1
                first = False
    ws.freeze_panes = "A2"


def write_truncation_diag_sheet(ws, sheets_before: dict[str, pd.DataFrame],
                                 sheets_after: dict[str, pd.DataFrame],
                                 cap: int, cap_desc: str) -> None:
    ws.title = "Truncation Diagnostics"
    headers  = ["Sheet", "News Type", "n", "Orig Median", "Post-Trunc Median",
                "Cap", "% Rows at Cap", "Cap Source"]
    _widths(ws, [42, 12, 8, 14, 18, 8, 16, 50])
    for ci, h in enumerate(headers, 1):
        _hdr(ws, 1, ci, h)
    ws.row_dimensions[1].height = 28

    ri = 2
    for sheet_name in sheets_before:
        bdf = sheets_before[sheet_name]
        adf = sheets_after.get(sheet_name, pd.DataFrame())
        for nt in ["HR", "AI-R", "HF", "AI-F"]:
            b_grp = bdf[bdf["news_type"] == nt]["token_count"]
            a_grp = adf[adf["news_type"] == nt]["token_count"] if len(adf) else pd.Series([], dtype=int)
            if len(b_grp) == 0:
                continue
            pct   = (a_grp == cap).mean() if len(a_grp) else float("nan")
            fill  = NT_FILLS.get(nt, ALT_ROW)
            p_fill = RED_FILL if pct > 0.5 else (YELLOW_FILL if pct > 0.2 else GREEN_FILL)
            _cell(ws, ri, 1, sheet_name,                  bold=True, align="left")
            _cell(ws, ri, 2, nt,                           fill_hex=fill)
            _cell(ws, ri, 3, len(b_grp),                   fill_hex=fill)
            _cell(ws, ri, 4, round(b_grp.median(), 1),     fill_hex=fill, fmt="0.0")
            _cell(ws, ri, 5, round(a_grp.median(), 1) if len(a_grp) else "—",
                  fill_hex=fill, fmt="0.0")
            _cell(ws, ri, 6, cap,                          fill_hex=YELLOW_FILL, bold=True)
            _cell(ws, ri, 7, pct,                          fill_hex=p_fill, fmt="0.0%")
            _cell(ws, ri, 8, cap_desc,                     align="left")
            ri += 1

    ri += 2
    for note in [
        "Cap = median token count of the category with the lowest overall median across all sheets/splits.",
        "Green = <20% rows truncated  |  Yellow = 20–50%  |  Red = >50%",
        "Truncation preserves [CLS] and [SEP] tokens; content tokens = cap − 2.",
    ]:
        c = ws.cell(row=ri, column=1, value=note)
        c.font = Font(name="Arial", italic=True, size=9, color="595959")
        ri += 1

    ws.freeze_panes = "A2"


# ── Main ──────────────────────────────────────────────────────────────────────

def load_sheets(path: str) -> dict[str, pd.DataFrame]:
    xl      = pd.ExcelFile(path)
    skip    = {"summary", "simulation summary", "truncation diagnostics",
               "test accuracy", "token stats (before)", "token stats (after)"}
    sheets  = {}
    for name in xl.sheet_names:
        if name.lower() in skip:
            continue
        df = xl.parse(name)
        df.columns = df.columns.str.strip()

        # Normalise key columns case-insensitively
        col_map = {c.lower(): c for c in df.columns}
        for key in ("news_type", "split", "label"):
            if key in col_map and col_map[key] != key:
                df.rename(columns={col_map[key]: key}, inplace=True)

        if "news_type" not in df.columns or "split" not in df.columns:
            print(f"  [SKIP] '{name}' — missing news_type or split.")
            continue

        df["news_type"] = df["news_type"].astype(str).str.upper().str.strip()
        df["split"]     = df["split"].astype(str).str.lower().str.strip()
        df = df[df["split"].isin({"train", "test", "val"})].copy()

        if len(df) == 0:
            print(f"  [SKIP] '{name}' — no usable rows.")
            continue

        sheets[name] = df

    return sheets


def print_summary(before: list[dict], after: list[dict], cap: int) -> None:
    print(f"\n{'═' * 72}")
    print(f"  RESULTS  (global cap = {cap} tokens)")
    print(f"{'═' * 72}")
    print(f"  {'Sheet':<40} {'Before':>10} {'After':>10} {'Δ':>8}")
    print(f"  {'─' * 70}")
    b_map = {r["sheet"]: r for r in before}
    a_map = {r["sheet"]: r for r in after}
    for sheet in b_map:
        if sheet not in a_map:
            continue
        b = b_map[sheet]["test_acc"]
        a = a_map[sheet]["test_acc"]
        arrow = "▼" if (a - b) < -0.01 else ("▲" if (a - b) > 0.01 else "~")
        print(f"  {sheet:<40} {b:>10.4f} {a:>10.4f} {a-b:>+8.4f} {arrow}")
    print(f"  {'─' * 70}")
    b_accs = [b_map[s]["test_acc"] for s in b_map if s in a_map]
    a_accs = [a_map[s]["test_acc"] for s in b_map if s in a_map]
    print(f"  {'MEAN':<40} {np.mean(b_accs):>10.4f} {np.mean(a_accs):>10.4f} "
          f"{np.mean(a_accs)-np.mean(b_accs):>+8.4f}")
    print()


def pick_mode() -> str:
    """Interactive mode picker shown at startup."""
    print()
    print("  ┌─────────────────────────────────────────────────────┐")
    print("  │         Length bias removal — choose approach        │")
    print("  ├─────────────────────────────────────────────────────┤")
    print("  │  1  Truncate only                                    │")
    print("  │     Cap all articles at HF mean. No resampling.     │")
    print("  │     Simplest — good for seeing raw truncation effect │")
    print("  │                                                      │")
    print("  │  2  Truncate + Simulate                              │")
    print("  │     Cap first, then resample AI-F→HF, AI-R→HR.      │")
    print("  │     Stronger bias removal. More moving parts.        │")
    print("  └─────────────────────────────────────────────────────┘")
    print()
    while True:
        choice = input("  Enter 1 or 2: ").strip()
        if choice in ("1", "2"):
            return choice
        print("  Please enter 1 or 2.")


def main() -> None:
    paths = sys.argv[1:]
    path  = paths[0] if paths else input("Path to stratified_dataset.xlsx: ").strip()

    mode = pick_mode()
    mode_label = "Truncate only" if mode == "1" else "Truncate + Simulate"
    print(f"\n  Mode: {mode_label}")

    print(f"\n  Loading '{path}' …")
    raw_sheets = load_sheets(path)
    print(f"  Found {len(raw_sheets)} valid sheet(s).")

    # ── Step 1: Tokenize ──────────────────────────────────────────────────────
    tok_sheets = step_tokenize(raw_sheets)

    # Save natural (pre-truncation) token counts for use as simulation reference
    natural_clf: dict[str, pd.DataFrame] = {
        name: df[["token_count", "label", "split", "news_type"]].copy()
        for name, df in tok_sheets.items()
    }

    # ── Step 2: Truncate (train only, mean-based cap) ─────────────────────────
    trunc_sheets, cap, cap_desc = step_truncate(tok_sheets)

    # Post-truncation token counts for train; test stays natural
    clf_orig: dict[str, pd.DataFrame] = {
        name: df[["token_count", "label", "split", "news_type"]].copy()
        for name, df in trunc_sheets.items()
    }

    # ── Step 3: Simulate (optional) ───────────────────────────────────────────
    if mode == "2":
        # Use natural stats as reference so test simulation is not contaminated
        # by truncated train distributions
        sim_sheets = step_simulate(clf_orig, natural_clf)
        after_source = sim_sheets
        after_label  = "After (truncate + simulate)"
    else:
        # No simulation — "after" is just the truncated dataset itself
        print("\n  [3/4] SIMULATE  →  skipped (truncate-only mode)")
        after_source = clf_orig
        after_label  = "After (truncate only)"

    # ── Step 4: Classify ──────────────────────────────────────────────────────
    before_results, after_results = step_classify(clf_orig, after_source)

    # ── Console summary ───────────────────────────────────────────────────────
    print_summary(before_results, after_results, cap)

    # ── Write Excel output ────────────────────────────────────────────────────
    wb = Workbook()
    wb.remove(wb.active)

    write_accuracy_sheet(
        wb.create_sheet("Test Accuracy"),
        before_results, after_results, cap,
        f"[{mode_label}] {cap_desc}",
    )
    write_token_stats_sheet(
        wb.create_sheet("Token Stats (Before)"), clf_orig,     "Token Stats (Before)"
    )
    write_token_stats_sheet(
        wb.create_sheet("Token Stats (After)"),  after_source, "Token Stats (After)"
    )
    write_truncation_diag_sheet(
        wb.create_sheet("Truncation Diagnostics"),
        natural_clf, clf_orig, cap, cap_desc,
    )

    out_path = pathlib.Path("length_pipeline_results.xlsx")
    wb.save(str(out_path))
    print(f"  ✓ Saved → {out_path}")
    print(f"  Mode used: {mode_label}")


if __name__ == "__main__":
    main()