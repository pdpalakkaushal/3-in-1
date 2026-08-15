#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
GridSync - Data Automation Tool (Python version)
=================================================
Teen kaam karta hai:
  1) Merge & Format  - multiple files ko aapke standard format me merge karta hai
  2) Compare Files   - do files ke beech difference nikalta hai
  3) SKU Match       - client SKU ko aapke system ke SKU se match karta hai (naam alag ho tab bhi)

Setup (ek baar):
  pip install pandas openpyxl

Chalane ke liye:
  python gridsync.py

Sab kuch aapke computer par hi chalta hai, koi internet ya data upload nahi hota.
"""

import sys
import os
import difflib
import re

try:
    import pandas as pd
except ImportError:
    print("pandas install nahi hai. Ye command chalayein: pip install pandas openpyxl")
    sys.exit(1)


# ---------------------------------------------------------------------------
# Utils
# ---------------------------------------------------------------------------

def normalize(value):
    """Text ko compare karne layak banata hai: lowercase, extra symbols/spaces hataata hai."""
    s = "" if value is None else str(value)
    s = s.lower().strip()
    s = re.sub(r"[^a-z0-9]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def similarity(a, b):
    """0 se 1 ke beech score - kitna similar hain do strings (naam alag hote hue bhi)."""
    na, nb = normalize(a), normalize(b)
    if not na and not nb:
        return 1.0
    if not na or not nb:
        return 0.0
    if na == nb:
        return 1.0
    return difflib.SequenceMatcher(None, na, nb).ratio()


def read_any_file(path):
    """xlsx, xls, ya csv - kisi bhi format ki file padh leta hai."""
    ext = os.path.splitext(path)[1].lower()
    if ext == ".csv":
        return pd.read_csv(path, dtype=str, keep_default_na=False)
    else:
        return pd.read_excel(path, dtype=str, keep_default_na=False)


def save_workbook(sheets, filename):
    """Dictionary of {sheet_name: DataFrame} ko ek Excel file me save karta hai."""
    with pd.ExcelWriter(filename, engine="openpyxl") as writer:
        for name, df in sheets.items():
            safe_name = str(name)[:31] if name else "Sheet1"
            if df is None or df.empty:
                df = pd.DataFrame([{"Info": "Koi row nahi mili"}])
            df.to_excel(writer, sheet_name=safe_name, index=False)
    print(f"\n✅ Saved: {filename}\n")


def ask_path(prompt):
    while True:
        path = input(prompt).strip().strip('"')
        if os.path.isfile(path):
            return path
        print("   Ye file nahi mili, dobara path daalein (ya poora path copy-paste karein).")


def ask_paths_multi(prompt):
    print(prompt)
    print("   (Ek file path likhein aur Enter dabayein. Khali line ke saath khatam karein.)")
    paths = []
    while True:
        p = input(f"   File {len(paths) + 1} (ya khali Enter khatam karne ke liye): ").strip().strip('"')
        if not p:
            if paths:
                break
            print("   Kam se kam ek file to daalni hogi.")
            continue
        if not os.path.isfile(p):
            print("   Ye file nahi mili, dobara try karein.")
            continue
        paths.append(p)
    return paths


def pick_column(columns, prompt="Column choose karein"):
    print(f"\n{prompt}:")
    for i, c in enumerate(columns, 1):
        print(f"   {i}. {c}")
    while True:
        choice = input("   Number daalein (ya khali Enter skip karne ke liye): ").strip()
        if choice == "":
            return None
        if choice.isdigit() and 1 <= int(choice) <= len(columns):
            return columns[int(choice) - 1]
        print("   Sahi number daalein.")


# ---------------------------------------------------------------------------
# Tool 1: Merge & Format
# ---------------------------------------------------------------------------

def tool_merge():
    print("\n=== Merge & Format ===")
    print("Pehle apna standard format define karein (columns, comma se separate karein).")
    default_fields = "SKU Code, SKU Name, Quantity, Date"
    raw = input(f"Standard columns [{default_fields}]: ").strip()
    fields = [f.strip() for f in (raw or default_fields).split(",") if f.strip()]

    paths = ask_paths_multi("\nAb wo files daalein jinhe merge karna hai:")

    all_rows = []
    for path in paths:
        df = read_any_file(path)
        columns = list(df.columns)

        # auto-guess mapping
        mapping = {}
        for field in fields:
            best_col, best_score = None, 0.0
            for c in columns:
                sc = similarity(field, c)
                if sc > best_score:
                    best_score, best_col = sc, c
            mapping[field] = best_col if best_score >= 0.45 else None

        print(f"\n--- {os.path.basename(path)} ({len(df)} rows) ---")
        print("Auto-detected mapping (galat lage to sahi karein):")
        for field in fields:
            guess = mapping[field] or "(nahi mila)"
            print(f"   {field}  ->  {guess}")
        change = input("Kisi column ki mapping badalni hai? (y/N): ").strip().lower()
        if change == "y":
            for field in fields:
                col = pick_column(columns, f"'{field}' ke liye source column")
                mapping[field] = col

        for _, row in df.iterrows():
            out = {}
            for field in fields:
                col = mapping[field]
                out[field] = row[col] if col else ""
            out["Source File"] = os.path.basename(path)
            all_rows.append(out)

    merged_df = pd.DataFrame(all_rows, columns=fields + ["Source File"])
    print(f"\nTotal merged rows: {len(merged_df)}")
    out_name = input("Output file ka naam [merged-standard-format.xlsx]: ").strip() or "merged-standard-format.xlsx"
    save_workbook({"Merged Data": merged_df}, out_name)


# ---------------------------------------------------------------------------
# Tool 2: Compare Files
# ---------------------------------------------------------------------------

def tool_compare():
    print("\n=== Compare Files ===")
    path_a = ask_path("Pehli file ka path: ")
    df_a = read_any_file(path_a)
    key_a = pick_column(list(df_a.columns), "File A me key column (jaise SKU/ID)")

    path_b = ask_path("\nDoosri file ka path: ")
    df_b = read_any_file(path_b)
    key_b = pick_column(list(df_b.columns), "File B me key column")

    if not key_a or not key_b:
        print("Key column zaroori hai, dobara try karein.")
        return

    map_a = {normalize(v): idx for idx, v in df_a[key_a].items()}
    map_b = {normalize(v): idx for idx, v in df_b[key_b].items()}

    common_cols = [c for c in df_a.columns if c in df_b.columns]

    changed_rows = []
    only_a_rows = []
    only_b_rows = []

    for norm_key, idx_a in map_a.items():
        if norm_key not in map_b:
            only_a_rows.append(df_a.loc[idx_a].to_dict())
            continue
        idx_b = map_b[norm_key]
        row_a, row_b = df_a.loc[idx_a], df_b.loc[idx_b]
        for col in common_cols:
            va, vb = str(row_a[col]).strip(), str(row_b[col]).strip()
            if va != vb:
                changed_rows.append({
                    "Key": row_a[key_a],
                    "Column": col,
                    f"Value in {os.path.basename(path_a)}": va,
                    f"Value in {os.path.basename(path_b)}": vb,
                })

    for norm_key, idx_b in map_b.items():
        if norm_key not in map_a:
            only_b_rows.append(df_b.loc[idx_b].to_dict())

    print(f"\nChanged cells: {len(changed_rows)}")
    print(f"Sirf {os.path.basename(path_a)} me: {len(only_a_rows)}")
    print(f"Sirf {os.path.basename(path_b)} me: {len(only_b_rows)}")

    out_name = input("\nOutput file ka naam [comparison-report.xlsx]: ").strip() or "comparison-report.xlsx"
    save_workbook({
        "Changed": pd.DataFrame(changed_rows),
        f"Only in {os.path.basename(path_a)}"[:31]: pd.DataFrame(only_a_rows),
        f"Only in {os.path.basename(path_b)}"[:31]: pd.DataFrame(only_b_rows),
    }, out_name)


# ---------------------------------------------------------------------------
# Tool 3: SKU Match
# ---------------------------------------------------------------------------

def tool_match():
    print("\n=== SKU Match ===")
    client_path = ask_path("Client ki SKU file ka path: ")
    df_client = read_any_file(client_path)
    client_cols = list(df_client.columns)
    client_sku_col = pick_column(client_cols, "Client file me SKU/code column")
    guess_pack = next((c for c in client_cols if re.search(r"pack ?shot|image|img|photo", c, re.I)), None)
    print(f"\nPackshot/image column guess: {guess_pack or '(nahi mila)'}")
    change = input("Packshot column badalna hai? (y/N): ").strip().lower()
    packshot_col = pick_column(client_cols, "Packshot column") if change == "y" else guess_pack

    internal_path = ask_path("\nHamare system ki SKU file ka path: ")
    df_internal = read_any_file(internal_path)
    internal_sku_col = pick_column(list(df_internal.columns), "System file me SKU/code column")

    if not client_sku_col or not internal_sku_col:
        print("SKU column zaroori hai, dobara try karein.")
        return

    threshold_raw = input("Fuzzy match sensitivity 50-95 (default 75): ").strip()
    threshold = (int(threshold_raw) if threshold_raw.isdigit() else 75) / 100.0

    exact_map = {}
    for idx, val in df_internal[internal_sku_col].items():
        exact_map[normalize(val)] = idx

    results = []
    counts = {"Exact": 0, "Fuzzy": 0, "Unmatched": 0}

    for _, row in df_client.iterrows():
        client_val = row[client_sku_col]
        norm = normalize(client_val)
        match_type, matched_val, confidence = "Unmatched", "", 0.0

        if norm in exact_map:
            matched_val = df_internal.loc[exact_map[norm], internal_sku_col]
            match_type, confidence = "Exact", 1.0
        else:
            best_score, best_val = 0.0, None
            for _, irow in df_internal.iterrows():
                sc = similarity(client_val, irow[internal_sku_col])
                if sc > best_score:
                    best_score, best_val = sc, irow[internal_sku_col]
            if best_val is not None and best_score >= threshold:
                matched_val, match_type, confidence = best_val, "Fuzzy", best_score

        counts[match_type] += 1
        rec = {
            "Client SKU": client_val,
            "Matched SKU (Hamara System)": matched_val,
            "Match Type": match_type,
            "Confidence %": round(confidence * 100),
        }
        if packshot_col:
            rec["Packshot"] = row[packshot_col]
        results.append(rec)

    print(f"\n✔ Exact: {counts['Exact']}   ⚠ Fuzzy: {counts['Fuzzy']}   ✕ Unmatched: {counts['Unmatched']}")

    out_name = input("\nOutput file ka naam [sku-match-report.xlsx]: ").strip() or "sku-match-report.xlsx"
    save_workbook({"SKU Match": pd.DataFrame(results)}, out_name)


# ---------------------------------------------------------------------------
# Main menu
# ---------------------------------------------------------------------------

def main():
    while True:
        print("\n" + "=" * 50)
        print("  GridSync — Data Automation Tool")
        print("=" * 50)
        print("1. Merge & Format (multiple files -> ek standard file)")
        print("2. Compare Files (do files ka difference)")
        print("3. SKU Match (client SKU <-> system SKU, packshot sahit)")
        print("4. Exit")
        choice = input("\nChoice daalein (1-4): ").strip()

        try:
            if choice == "1":
                tool_merge()
            elif choice == "2":
                tool_compare()
            elif choice == "3":
                tool_match()
            elif choice == "4":
                print("Bye!")
                break
            else:
                print("Sahi option choose karein (1-4).")
        except Exception as e:
            print(f"\n⚠ Kuch error aaya: {e}")
            print("File format ya column selection check karke dobara try karein.")


if __name__ == "__main__":
    main()
