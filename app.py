import streamlit as st
import pandas as pd
import difflib
import re
import io
import os

st.set_page_config(page_title="GridSync", page_icon="📊", layout="wide")

# =====================================================================
# Utils
# =====================================================================

def normalize(value):
    s = "" if value is None else str(value)
    s = s.lower().strip()
    s = re.sub(r"[^a-z0-9]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def similarity(a, b):
    na, nb = normalize(a), normalize(b)
    if not na and not nb:
        return 1.0
    if not na or not nb:
        return 0.0
    if na == nb:
        return 1.0
    return difflib.SequenceMatcher(None, na, nb).ratio()


def read_uploaded_file(uploaded_file):
    ext = os.path.splitext(uploaded_file.name)[1].lower()
    if ext == ".csv":
        return pd.read_csv(uploaded_file, dtype=str, keep_default_na=False)
    return pd.read_excel(uploaded_file, dtype=str, keep_default_na=False)


def to_excel_bytes(sheets: dict):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for name, df in sheets.items():
            safe_name = str(name)[:31] if name else "Sheet1"
            if df is None or df.empty:
                df = pd.DataFrame([{"Info": "Koi row nahi mili"}])
            df.to_excel(writer, sheet_name=safe_name, index=False)
    return output.getvalue()


EXCEL_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

# =====================================================================
# Sidebar navigation
# =====================================================================

st.sidebar.title("📊 GridSync")
st.sidebar.caption("Data daaliye, baaki khud ho jaayega")
tool = st.sidebar.radio("Tool choose karein", ["Merge & Format", "Compare Files", "SKU Match"])

# =====================================================================
# TOOL 1: Merge & Format
# =====================================================================

if tool == "Merge & Format":
    st.header("Merge & Format")
    st.caption("Multiple call cycle files ko apne standard format me merge karein.")

    if "std_fields" not in st.session_state:
        st.session_state.std_fields = ["SKU Code", "SKU Name", "Quantity", "Date"]

    st.subheader("1. Standard format define karein")
    c1, c2 = st.columns([4, 1])
    new_field = c1.text_input("Naya column", placeholder="jaise Client Name", label_visibility="collapsed")
    if c2.button("+ Add field") and new_field.strip():
        if new_field.strip() not in st.session_state.std_fields:
            st.session_state.std_fields.append(new_field.strip())
        st.rerun()

    if st.session_state.std_fields:
        chip_cols = st.columns(len(st.session_state.std_fields))
        for i, f in enumerate(list(st.session_state.std_fields)):
            if chip_cols[i].button(f"✕ {f}", key=f"remove_{f}"):
                st.session_state.std_fields.remove(f)
                st.rerun()

    st.subheader("2. Files upload karein")
    uploaded_files = st.file_uploader(
        "Call cycle files", type=["xlsx", "xls", "csv"], accept_multiple_files=True
    )

    if uploaded_files:
        st.subheader("3. Column mapping check karein")
        mappings = {}
        for uf in uploaded_files:
            df = read_uploaded_file(uf)
            st.markdown(f"**📄 {uf.name}** ({len(df)} rows)")
            file_map = {}
            map_cols = st.columns(2)
            for i, field in enumerate(st.session_state.std_fields):
                best_col, best_score = None, 0.0
                for c in df.columns:
                    sc = similarity(field, c)
                    if sc > best_score:
                        best_score, best_col = sc, c
                options = ["— khali chhodein —"] + list(df.columns)
                default_idx = options.index(best_col) if best_score >= 0.45 and best_col else 0
                with map_cols[i % 2]:
                    choice = st.selectbox(field, options, index=default_idx, key=f"map_{uf.name}_{field}")
                file_map[field] = None if choice == "— khali chhodein —" else choice
            mappings[uf.name] = (df, file_map)
            st.divider()

        if st.button("Merge karein", type="primary"):
            all_rows = []
            for fname, (df, file_map) in mappings.items():
                for _, row in df.iterrows():
                    out = {field: (row[col] if col else "") for field, col in file_map.items()}
                    out["Source File"] = fname
                    all_rows.append(out)
            st.session_state.merged_df = pd.DataFrame(
                all_rows, columns=st.session_state.std_fields + ["Source File"]
            )

    if "merged_df" in st.session_state:
        merged_df = st.session_state.merged_df
        st.success(f"✅ Merged output ({len(merged_df)} rows)")
        st.dataframe(merged_df.head(50), use_container_width=True)
        st.download_button(
            "⬇ Excel download",
            data=to_excel_bytes({"Merged Data": merged_df}),
            file_name="merged-standard-format.xlsx",
            mime=EXCEL_MIME,
        )

# =====================================================================
# TOOL 2: Compare Files
# =====================================================================

elif tool == "Compare Files":
    st.header("Compare Files")
    st.caption("Do files ke beech difference nikalein.")

    c1, c2 = st.columns(2)
    df_a = df_b = None
    key_a = key_b = None

    with c1:
        file_a = st.file_uploader("Pehli file", type=["xlsx", "xls", "csv"], key="file_a")
        if file_a:
            df_a = read_uploaded_file(file_a)
            key_a = st.selectbox("File A key column", df_a.columns, key="key_a")
    with c2:
        file_b = st.file_uploader("Doosri file", type=["xlsx", "xls", "csv"], key="file_b")
        if file_b:
            df_b = read_uploaded_file(file_b)
            key_b = st.selectbox("File B key column", df_b.columns, key="key_b")

    if st.button("🔀 Compare karein", type="primary", disabled=not (file_a and file_b)):
        map_a = {normalize(v): idx for idx, v in df_a[key_a].items()}
        map_b = {normalize(v): idx for idx, v in df_b[key_b].items()}
        common_cols = [c for c in df_a.columns if c in df_b.columns]

        changed_rows, only_a_rows, only_b_rows = [], [], []
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
                        f"Value in {file_a.name}": va,
                        f"Value in {file_b.name}": vb,
                    })
        for norm_key, idx_b in map_b.items():
            if norm_key not in map_a:
                only_b_rows.append(df_b.loc[idx_b].to_dict())

        st.session_state.compare_result = {
            "changed": pd.DataFrame(changed_rows),
            "only_a": pd.DataFrame(only_a_rows),
            "only_b": pd.DataFrame(only_b_rows),
            "name_a": file_a.name,
            "name_b": file_b.name,
        }

    if "compare_result" in st.session_state:
        res = st.session_state.compare_result
        t1, t2, t3 = st.tabs([
            f"Changed ({len(res['changed'])})",
            f"Sirf {res['name_a']} me ({len(res['only_a'])})",
            f"Sirf {res['name_b']} me ({len(res['only_b'])})",
        ])
        with t1:
            st.dataframe(res["changed"], use_container_width=True)
        with t2:
            st.dataframe(res["only_a"], use_container_width=True)
        with t3:
            st.dataframe(res["only_b"], use_container_width=True)

        st.download_button(
            "⬇ Report download",
            data=to_excel_bytes({
                "Changed": res["changed"],
                f"Only in {res['name_a']}"[:31]: res["only_a"],
                f"Only in {res['name_b']}"[:31]: res["only_b"],
            }),
            file_name="comparison-report.xlsx",
            mime=EXCEL_MIME,
        )

# =====================================================================
# TOOL 3: SKU Match
# =====================================================================

else:
    st.header("SKU Match")
    st.caption("Client SKU ko aapke system ke SKU se match karein — naam alag hote hue bhi.")

    c1, c2 = st.columns(2)
    df_client = df_internal = None
    client_sku_col = client_pack_col = internal_sku_col = None

    with c1:
        client_file = st.file_uploader("Client ki SKU file", type=["xlsx", "xls", "csv"], key="client_file")
        if client_file:
            df_client = read_uploaded_file(client_file)
            guess_sku = next((c for c in df_client.columns if re.search(r"sku|code|item", c, re.I)), df_client.columns[0])
            guess_pack = next((c for c in df_client.columns if re.search(r"pack ?shot|image|img|photo", c, re.I)), None)
            client_sku_col = st.selectbox(
                "SKU column", df_client.columns, index=list(df_client.columns).index(guess_sku), key="client_sku_col"
            )
            pack_options = ["— nahi hai —"] + list(df_client.columns)
            default_pack_idx = pack_options.index(guess_pack) if guess_pack else 0
            pack_choice = st.selectbox("Packshot column", pack_options, index=default_pack_idx, key="client_pack_col")
            client_pack_col = None if pack_choice == "— nahi hai —" else pack_choice

    with c2:
        internal_file = st.file_uploader("Hamare system ki SKU file", type=["xlsx", "xls", "csv"], key="internal_file")
        if internal_file:
            df_internal = read_uploaded_file(internal_file)
            guess_sku2 = next((c for c in df_internal.columns if re.search(r"sku|code|item", c, re.I)), df_internal.columns[0])
            internal_sku_col = st.selectbox(
                "SKU column", df_internal.columns, index=list(df_internal.columns).index(guess_sku2), key="internal_sku_col"
            )

    threshold = st.slider("Fuzzy match sensitivity", 50, 95, 75, step=5, format="%d%%") / 100.0

    if st.button("🔎 Match karein", type="primary", disabled=not (client_file and internal_file)):
        exact_map = {normalize(v): idx for idx, v in df_internal[internal_sku_col].items()}
        results = []
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

            rec = {
                "Client SKU": client_val,
                "Matched SKU (Hamara System)": matched_val,
                "Match Type": match_type,
                "Confidence %": round(confidence * 100),
            }
            if client_pack_col:
                rec["Packshot"] = row[client_pack_col]
            results.append(rec)

        st.session_state.match_results = pd.DataFrame(results)

    if "match_results" in st.session_state:
        results_df = st.session_state.match_results
        counts = results_df["Match Type"].value_counts().to_dict()
        m1, m2, m3 = st.columns(3)
        m1.metric("✔ Exact", counts.get("Exact", 0))
        m2.metric("⚠ Fuzzy", counts.get("Fuzzy", 0))
        m3.metric("✕ Unmatched", counts.get("Unmatched", 0))

        view = st.radio("View", ["Sabhi", "Fuzzy", "Unmatched"], horizontal=True)
        if view == "Fuzzy":
            show_df = results_df[results_df["Match Type"] == "Fuzzy"]
        elif view == "Unmatched":
            show_df = results_df[results_df["Match Type"] == "Unmatched"]
        else:
            show_df = results_df
        st.dataframe(show_df, use_container_width=True)

        st.download_button(
            "⬇ Report download",
            data=to_excel_bytes({"SKU Match": results_df}),
            file_name="sku-match-report.xlsx",
            mime=EXCEL_MIME,
        )
