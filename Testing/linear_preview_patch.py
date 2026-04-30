"""
linear_preview_patch.py
-----------------------
Apply this patch to linear_preview.py to handle datasets with no val split.
Run once: python linear_preview_patch.py linear_preview.py
"""
import sys
import re

path = sys.argv[1] if len(sys.argv) > 1 else "linear_preview.py"
src  = open(path, encoding="utf-8").read()

# ── Patch 1: make run_sheet tolerate missing val split ────────────────────────
old = '''    train_df = df[df["split"] == "train"].reset_index(drop=True)
    val_df   = df[df["split"] == "val"].reset_index(drop=True)
    test_df  = df[df["split"] == "test"].reset_index(drop=True)

    clf = LogisticRegression(random_state=42, max_iter=1000)
    clf.fit(train_df[["token_count"]], train_df["label"])

    train_m = evaluate(clf, train_df)
    val_m   = evaluate(clf, val_df)
    test_m  = evaluate(clf, test_df)'''

new = '''    train_df = df[df["split"] == "train"].reset_index(drop=True)
    val_df   = df[df["split"] == "val"].reset_index(drop=True)
    test_df  = df[df["split"] == "test"].reset_index(drop=True)

    if len(train_df) == 0:
        raise ValueError(f"Sheet '{sheet_name}' has no train rows.")
    if len(test_df) == 0:
        raise ValueError(f"Sheet '{sheet_name}' has no test rows.")

    clf = LogisticRegression(random_state=42, max_iter=1000)
    clf.fit(train_df[["token_count"]], train_df["label"])

    train_m = evaluate(clf, train_df)
    # val is optional — fall back to a copy of test metrics if absent
    val_m   = evaluate(clf, val_df) if len(val_df) > 0 else evaluate(clf, test_df)
    test_m  = evaluate(clf, test_df)'''

if old in src:
    src = src.replace(old, new)
    print("  ✓ Patch 1 applied: val split is now optional.")
else:
    print("  [WARN] Patch 1 target not found — may already be patched or code differs.")

open(path, "w", encoding="utf-8").write(src)
print(f"  Saved → {path}")
