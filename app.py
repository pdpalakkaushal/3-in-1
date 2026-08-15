import streamlit as st
import pandas as pd
import difflib
import re
import io
import os
import base64
import openpyxl

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


def combine_cols(row, cols):
    """Multiple columns ko ek text me jod deta hai (matching ke liye)."""
    parts = [str(row[c]) for c in cols if c in row and pd.notna(row[c]) and str(row[c]).strip()]
    return " ".join(parts)


def first_nonempty(row, cols):
    """Diye gaye columns me se pehli non-empty value laata hai (packshot ke liye)."""
    for c in cols:
        if c in row and pd.notna(row[c]) and str(row[c]).strip():
            return row[c]
    return ""


def extract_embedded_images(file_bytes, filename):
    """
    Excel file ke andar seedhe chipki hui (embedded/pasted) images dhoondh kar
    {(row_index, column_name): data_uri} ke roop me laata hai.
    Ye SIRF System SKU Catalog aur SKU Match tabs me use hota hai — Merge/Compare
    is function ko chhoote tak nahi (unpar koi asar nahi).
    Sirf .xlsx files ke liye kaam karta hai — csv me embedded image possible nahi hoti.
    Assume karta hai header row 1 (excel) me hai — agar aisa nahi hai to skip ho jaata hai, error nahi aata.
    """
    if not file_bytes or not filename or not filename.lower().endswith(".xlsx"):
        return {}
    try:
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
        ws = wb.active
        headers = {}
        for cell in ws[1]:
            if cell.value:
                headers[cell.column] = str(cell.value).strip()

        images_map = {}
        for img in getattr(ws, "_images", []):
            try:
                anchor = img.anchor
                col_idx = anchor._from.col + 1
                row_idx = anchor._from.row + 1
                df_row = row_idx - 2  # excel row 1 = header -> df row 0 = excel row 2
                col_name = headers.get(col_idx)
                if col_name is None or df_row < 0:
                    continue
                img_bytes = img._data()
                b64 = base64.b64encode(img_bytes).decode("utf-8")
                images_map[(df_row, col_name)] = f"data:image/png;base64,{b64}"
            except Exception:
                continue
        return images_map
    except Exception:
        return {}


def enrich_with_embedded_images(df, file_bytes, filename, pack_cols):
    """
    Packshot column khali ho (link na ho) lekin usi cell me image seedhi paste/insert
    ki hui ho, to use dhoondh kar dataframe me bhar deta hai — taaki wo bhi match/preview ho sake.
    """
    if not pack_cols or not file_bytes:
        return df
    images_map = extract_embedded_images(file_bytes, filename)
    if not images_map:
        return df
    df = df.copy().reset_index(drop=True)
    for col in pack_cols:
        if col not in df.columns:
            continue
        for i in df.index:
            current_val = df.at[i, col]
            if not str(current_val).strip():
                data_uri = images_map.get((i, col))
                if data_uri:
                    df.at[i, col] = data_uri
    return df


EXCEL_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
IMG_PATTERN = re.compile(r"image|packshot|pack ?shot|photo|img|picture", re.I)
GROUP_PATTERN = re.compile(r"group|sku name|class|item name", re.I)
CLASS_PATTERN = re.compile(r"class|sku|item", re.I)

# =====================================================================
# Sidebar navigation
# =====================================================================

st.sidebar.title("📊 GridSync")
st.sidebar.caption("Data daaliye, baaki khud ho jaayega")
tool = st.sidebar.radio(
    "Tool choose karein",
    ["Merge & Format", "Compare Files", "System SKU Catalog", "SKU Match"],
)

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
# TOOL 3: System SKU Catalog  (NEW)
# =====================================================================

elif tool == "System SKU Catalog":
    st.header("System SKU Catalog")
    st.caption(
        "Hamare system ki SKU/Group/Class sheet ek baar yahan upload karein — "
        "iska data 'SKU Match' tab me automatically reuse hoga, dobara upload nahi karna padega."
    )

    catalog_file = st.file_uploader(
        "System ki sheet upload karein (jaise: category, group_name, class_name, image link)",
        type=["xlsx", "xls", "csv"],
        key="catalog_file",
    )
    if catalog_file is not None:
        st.session_state.system_df = read_uploaded_file(catalog_file)
        st.session_state.system_file_name = catalog_file.name
        st.session_state.system_file_bytes = catalog_file.getvalue()

    if "system_df" in st.session_state:
        df_catalog = st.session_state.system_df
        cols = list(df_catalog.columns)

        st.subheader("Columns select karein")
        st.caption("Koi bhi column, kitne bhi select kar sakte hain — koi restriction nahi. Dropdown me click karke jitne chahiye utne jod dein.")
        c1, c2, c3 = st.columns(3)
        with c1:
            group_default = [c for c in cols if GROUP_PATTERN.search(c)][:1]
            system_group_cols = st.multiselect("Group column(s)", cols, default=group_default, key="sys_group_cols")
        with c2:
            class_default = [c for c in cols if CLASS_PATTERN.search(c) and c not in system_group_cols][:1]
            system_class_cols = st.multiselect("Class / SKU column(s)", cols, default=class_default, key="sys_class_cols")
        with c3:
            pack_default = [c for c in cols if IMG_PATTERN.search(c)]
            system_pack_cols = st.multiselect("Packshot column(s)", cols, default=pack_default, key="sys_pack_cols")

        # Packshot column me link na ho to, usi cell me paste ki hui image bhi dhoondh lein
        if system_pack_cols:
            df_catalog = enrich_with_embedded_images(
                df_catalog,
                st.session_state.get("system_file_bytes"),
                st.session_state.get("system_file_name", ""),
                system_pack_cols,
            )

        search = st.text_input("🔍 Search (naam se dhoondhein)")
        view_df = df_catalog.copy()
        if search:
            mask = view_df.apply(lambda r: search.lower() in " ".join(str(v) for v in r.values).lower(), axis=1)
            view_df = view_df[mask]

        display_df = view_df.copy()
        column_config = {}
        if system_pack_cols:
            display_df["Packshot"] = view_df.apply(lambda r: first_nonempty(r, system_pack_cols), axis=1)
            column_config["Packshot"] = st.column_config.ImageColumn("Packshot", width="small")

        st.dataframe(display_df, use_container_width=True, hide_index=True, column_config=column_config)
        st.caption(f"Total {len(df_catalog)} entries system me hain — {len(view_df)} dikh rahi hain.")
        st.info(
            "Packshot ke liye ye tool image ka **link (URL)** aur **.xlsx me seedhi paste/insert ki hui image** — dono support karta hai. "
            "Link ho to seedha uska preview aa jaayega; agar link nahi hai to cell me chipki image khud detect karke dikha di jaayegi."
        )
    else:
        st.info("Upar se apni system sheet upload karke shuru karein.")

# =====================================================================
# TOOL 4: SKU Match  (multi-select group/class/packshot + image preview)
# =====================================================================

else:
    st.header("SKU Match")
    st.caption("Client SKU ko aapke system ke SKU se match karein — naam alag hote hue bhi, packshot ke saath.")

    if "system_df" not in st.session_state:
        st.warning("Pehle **'System SKU Catalog'** tab me apni system sheet upload kar lein, phir yahan match karein.")
    else:
        df_internal = st.session_state.system_df
        internal_group_cols = st.session_state.get("sys_group_cols", [])
        internal_class_cols = st.session_state.get("sys_class_cols", [])
        internal_pack_cols = st.session_state.get("sys_pack_cols", [])
        internal_key_cols = internal_group_cols + internal_class_cols

        st.info(
            f"System catalog loaded: **{st.session_state.get('system_file_name','')}** "
            f"({len(df_internal)} rows). Columns badalne ho to 'System SKU Catalog' tab me jaayein."
        )

        if not internal_key_cols:
            st.warning("'System SKU Catalog' tab me pehle Group ya Class column select karein.")
        else:
            client_file = st.file_uploader("Client ki SKU file (jaise aapka screenshot wali sheet)", type=["xlsx", "xls", "csv"], key="client_file2")

            if client_file:
                df_client = read_uploaded_file(client_file)
                client_cols = list(df_client.columns)

                c1, c2 = st.columns(2)
                with c1:
                    group_default = [c for c in client_cols if GROUP_PATTERN.search(c)][:1]
                    client_group_cols = st.multiselect(
                        "Client Group / SKU Name column(s)", client_cols, default=group_default, key="client_group_cols"
                    )
                with c2:
                    pack_default = [c for c in client_cols if IMG_PATTERN.search(c)]
                    client_pack_cols = st.multiselect(
                        "Client Packshot column(s) (Front/Back — jitne bhi hon, sab select karein)",
                        client_cols, default=pack_default, key="client_pack_cols",
                    )

                threshold = st.slider("Fuzzy match sensitivity (naam thoda alag ho tab bhi match kare)", 50, 95, 75, step=5, format="%d%%") / 100.0

                if st.button("🔎 Match karein", type="primary", disabled=not client_group_cols):
                    # Packshot column me link na ho to, usi cell me paste ki hui image bhi dhoondh lein
                    # (system sheet aur client sheet dono ke liye)
                    df_internal_m = enrich_with_embedded_images(
                        df_internal,
                        st.session_state.get("system_file_bytes"),
                        st.session_state.get("system_file_name", ""),
                        internal_pack_cols,
                    )
                    df_client_m = enrich_with_embedded_images(
                        df_client, client_file.getvalue(), client_file.name, client_pack_cols
                    )

                    internal_records = []
                    for _, irow in df_internal_m.iterrows():
                        combined = combine_cols(irow, internal_key_cols)
                        pack_val = first_nonempty(irow, internal_pack_cols)
                        internal_records.append((combined, pack_val))
                    st.session_state.internal_records = internal_records

                    exact_map = {}
                    for combined, pack_val in internal_records:
                        norm = normalize(combined)
                        if norm not in exact_map:
                            exact_map[norm] = (combined, pack_val)

                    results = []
                    for _, crow in df_client_m.iterrows():
                        client_combined = combine_cols(crow, client_group_cols)
                        norm = normalize(client_combined)
                        match_type, matched_val, confidence, matched_pack = "Unmatched", "", 0.0, ""

                        if norm in exact_map:
                            matched_val, matched_pack = exact_map[norm]
                            match_type, confidence = "Exact", 1.0
                        else:
                            best_score, best_combined, best_pack = 0.0, None, ""
                            for combined, pack_val in internal_records:
                                sc = similarity(client_combined, combined)
                                if sc > best_score:
                                    best_score, best_combined, best_pack = sc, combined, pack_val
                            if best_combined is not None and best_score >= threshold:
                                matched_val, match_type, confidence, matched_pack = best_combined, "Fuzzy", best_score, best_pack

                        client_pack_val = first_nonempty(crow, client_pack_cols)

                        results.append({
                            "Client SKU/Group": client_combined,
                            "Client Packshot": client_pack_val,
                            "Matched SKU (System)": matched_val,
                            "Matched Packshot": matched_pack,
                            "Match Type": match_type,
                            "Confidence %": round(confidence * 100),
                        })

                    st.session_state.match_results = pd.DataFrame(results)

                if "match_results" in st.session_state:
                    results_df = st.session_state.match_results
                    counts = results_df["Match Type"].value_counts().to_dict()
                    total = len(results_df)
                    auto_matched = total - counts.get("Unmatched", 0)

                    m1, m2, m3, m4 = st.columns(4)
                    m1.metric("✔ Exact", counts.get("Exact", 0))
                    m2.metric("⚠ Fuzzy", counts.get("Fuzzy", 0))
                    m3.metric("✏️ Manual", counts.get("Manual", 0))
                    m4.metric("✕ Unmatched", counts.get("Unmatched", 0))
                    st.progress(
                        auto_matched / total if total else 0,
                        text=f"{auto_matched}/{total} SKU automatically match ho gaye ({round(100 * auto_matched / total) if total else 0}%)",
                    )

                    view = st.radio("View", ["Sabhi", "Fuzzy", "Unmatched"], horizontal=True)
                    if view == "Fuzzy":
                        show_df = results_df[results_df["Match Type"] == "Fuzzy"]
                    elif view == "Unmatched":
                        show_df = results_df[results_df["Match Type"] == "Unmatched"]
                    else:
                        show_df = results_df

                    st.dataframe(
                        show_df,
                        use_container_width=True,
                        hide_index=True,
                        column_config={
                            "Client Packshot": st.column_config.ImageColumn("Client Packshot", width="small"),
                            "Matched Packshot": st.column_config.ImageColumn("Matched Packshot", width="small"),
                        },
                    )

                    # Baaki bache unmatched items ko haath se (manually) theek karne ka option
                    unmatched_mask = results_df["Match Type"] == "Unmatched"
                    if unmatched_mask.any():
                        internal_records = st.session_state.get("internal_records", [])
                        options_list = ["— select —"] + [c for c, _ in internal_records if c]
                        with st.expander(f"✏️ {int(unmatched_mask.sum())} unmatched items ko manually fix karein"):
                            edit_df = results_df.loc[unmatched_mask, ["Client SKU/Group"]].copy()
                            edit_df["System me sahi match"] = "— select —"
                            edited = st.data_editor(
                                edit_df,
                                column_config={
                                    "System me sahi match": st.column_config.SelectboxColumn(
                                        "System me sahi match", options=options_list
                                    )
                                },
                                hide_index=True,
                                use_container_width=True,
                                key="manual_fix_editor",
                            )
                            if st.button("✅ Manual matches apply karein"):
                                pack_lookup = dict(internal_records)
                                for idx, row in edited.iterrows():
                                    chosen = row["System me sahi match"]
                                    if chosen and chosen != "— select —":
                                        results_df.loc[idx, "Matched SKU (System)"] = chosen
                                        results_df.loc[idx, "Matched Packshot"] = pack_lookup.get(chosen, "")
                                        results_df.loc[idx, "Match Type"] = "Manual"
                                        results_df.loc[idx, "Confidence %"] = 100
                                st.session_state.match_results = results_df
                                st.rerun()

                    st.download_button(
                        "⬇ Report download",
                        data=to_excel_bytes({"SKU Match": results_df}),
                        file_name="sku-match-report.xlsx",
                        mime=EXCEL_MIME,
                    )
