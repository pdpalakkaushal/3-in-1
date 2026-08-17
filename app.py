import streamlit as st
import pandas as pd
import difflib
import re
import io
import os
import base64
from io import BytesIO

from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter
from PIL import Image as PILImage

st.set_page_config(page_title="GridSync", page_icon="📊", layout="wide")

# =====================================================================
# Shared utils (used across tools)
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
# SKU Match specific utils (kept separate so nothing else is touched)
# =====================================================================

IMG_PATTERN = re.compile(r"image|packshot|pack ?shot|photo|img|picture|visual", re.I)

# Tiered keyword patterns: checked in priority order, most specific first.
NAME_PRIORITY_PATTERNS = [
    re.compile(r"group.?name", re.I),
    re.compile(r"class.?name", re.I),
    re.compile(r"sku.?name", re.I),
    re.compile(r"product.?name", re.I),
    re.compile(r"item.?name", re.I),
    re.compile(r"\bgroup\b", re.I),
    re.compile(r"\bclass\b", re.I),
    re.compile(r"\bsku\b", re.I),
    re.compile(r"\bitem\b", re.I),
]


def combine_cols(row, cols):
    """Multiple columns ko ek text me jod deta hai (matching ke liye)."""
    parts = [str(row[c]) for c in cols if c in row and pd.notna(row[c]) and str(row[c]).strip()]
    return " ".join(parts)


def first_nonempty(row, cols):
    for c in cols:
        if c in row and pd.notna(row[c]) and str(row[c]).strip():
            return row[c]
    return ""


def get_file_bytes(uploaded_file):
    uploaded_file.seek(0)
    data = uploaded_file.read()
    uploaded_file.seek(0)
    return data


def read_df_from_bytes(file_bytes, filename):
    ext = os.path.splitext(filename)[1].lower()
    if ext == ".csv":
        return pd.read_csv(BytesIO(file_bytes), dtype=str, keep_default_na=False)
    return pd.read_excel(BytesIO(file_bytes), dtype=str, keep_default_na=False)


def extract_embedded_images(file_bytes, filename):
    """
    xlsx ke andar seedhi chipki hui (embedded) images nikaal kar
    {(data_row_index, column_name): image_bytes} banata hai.
    .xls aur .csv me embedded images support nahi hoti, unke liye khali dict.
    """
    ext = os.path.splitext(filename)[1].lower()
    if ext != ".xlsx":
        return {}
    try:
        wb = load_workbook(BytesIO(file_bytes))
        ws = wb.active
        header_row = 1
        col_name_by_idx = {}
        for cell in ws[header_row]:
            if cell.value is not None:
                col_name_by_idx[cell.column] = str(cell.value).strip()

        images_map = {}
        for img in getattr(ws, "_images", []):
            try:
                anchor = img.anchor
                from_marker = anchor._from
                col_1idx = from_marker.col + 1
                row_0idx = from_marker.row
                col_name = col_name_by_idx.get(col_1idx)
                if not col_name:
                    continue
                data_row_idx = row_0idx - header_row  # header sits at spreadsheet row 1 (0-idx row 0)
                if data_row_idx < 0:
                    continue
                images_map[(data_row_idx, col_name)] = img._data()
            except Exception:
                continue
        return images_map
    except Exception:
        return {}


def image_bytes_to_data_uri(img_bytes):
    try:
        pil_img = PILImage.open(BytesIO(img_bytes))
        fmt = (pil_img.format or "PNG").lower()
        mime = f"image/{fmt}"
    except Exception:
        mime = "image/png"
    b64 = base64.b64encode(img_bytes).decode("utf-8")
    return f"data:{mime};base64,{b64}"


def guess_name_columns(df, exclude_cols=None):
    """Product/SKU/Group naam wala column dhoondta hai — header keyword se, warna
    text ki length aur uniqueness dekh kar (bina kisi keyword ke bhi kaam karta hai)."""
    exclude_cols = exclude_cols or set()
    cols = [c for c in df.columns if c not in exclude_cols]

    for pattern in NAME_PRIORITY_PATTERNS:
        hits = [c for c in cols if pattern.search(c)]
        if hits:
            return hits[:2]

    best_col, best_score = None, -1.0
    for c in cols:
        vals = df[c].astype(str).str.strip()
        vals = vals[vals != ""]
        if len(vals) == 0:
            continue
        avg_len = vals.str.len().mean()
        uniqueness = vals.nunique() / max(len(vals), 1)
        score = avg_len * uniqueness
        if score > best_score:
            best_score, best_col = score, c
    return [best_col] if best_col else []


def guess_image_columns(df, embedded_cols_set):
    """Packshot columns dhoondta hai — embedded images, header keyword, ya URL text dekh kar."""
    keyword_cols = [c for c in df.columns if IMG_PATTERN.search(c)]
    url_cols = []
    for c in df.columns:
        sample = df[c].astype(str).str.strip().head(30)
        if sample.str.lower().str.startswith("http").any():
            url_cols.append(c)
    combined = list(dict.fromkeys(list(embedded_cols_set) + keyword_cols + url_cols))
    return combined


def get_packshot_value(row, row_idx, candidate_cols, embedded_map):
    """
    Ek row ke liye packshot nikalta hai — embedded image ho ya URL/text, dono handle karta hai.
    Returns: {"display": <table me dikhane wali cheez>, "bytes": <agar embedded to raw bytes>}
    """
    for c in candidate_cols:
        key = (row_idx, c)
        if key in embedded_map:
            return {"display": image_bytes_to_data_uri(embedded_map[key]), "bytes": embedded_map[key]}
    val = first_nonempty(row, candidate_cols)
    if val:
        return {"display": str(val), "bytes": None}
    return {"display": "", "bytes": None}


def build_match_excel(results_df, image_bytes_map):
    """
    SKU Match ka Excel report banata hai. Agar kisi row ka packshot embedded-source
    tha (raw bytes available hain), to usse seedhe Excel cell me bhi chipka deta hai.
    image_bytes_map: {"client": {row_idx: bytes}, "matched": {row_idx: bytes}}
    """
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        results_df.to_excel(writer, sheet_name="SKU Match", index=False)
        ws = writer.sheets["SKU Match"]

        headers = list(results_df.columns)
        col_letter = {h: get_column_letter(i + 1) for i, h in enumerate(headers)}

        for h in ["Client Packshot", "Matched Packshot"]:
            if h in col_letter:
                ws.column_dimensions[col_letter[h]].width = 16

        for row_idx in range(len(results_df)):
            excel_row = row_idx + 2  # header row = 1
            has_image = row_idx in image_bytes_map.get("client", {}) or row_idx in image_bytes_map.get("matched", {})
            if has_image:
                ws.row_dimensions[excel_row].height = 60
            for key, col_name in [("client", "Client Packshot"), ("matched", "Matched Packshot")]:
                bytes_map = image_bytes_map.get(key, {})
                if row_idx in bytes_map and col_name in col_letter:
                    try:
                        xl_img = XLImage(io.BytesIO(bytes_map[row_idx]))
                        xl_img.width, xl_img.height = 70, 70
                        ws.add_image(xl_img, f"{col_letter[col_name]}{excel_row}")
                    except Exception:
                        pass
    return output.getvalue()


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
# TOOL 1: Merge & Format   (UNCHANGED)
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
# TOOL 2: Compare Files   (UNCHANGED)
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
# TOOL 3: System SKU Catalog   (fully automatic, embedded-image aware)
# =====================================================================

elif tool == "System SKU Catalog":
    st.header("System SKU Catalog")
    st.caption(
        "Hamare system ki SKU/Group/Class sheet ek baar yahan upload karein — tool khud "
        "naam aur packshot wale columns pehchan lega. Ye data 'SKU Match' tab me apne aap reuse hoga."
    )

    catalog_file = st.file_uploader(
        "System ki sheet upload karein",
        type=["xlsx", "xls", "csv"],
        key="catalog_file",
    )
    if catalog_file is not None:
        file_bytes = get_file_bytes(catalog_file)
        st.session_state.system_df = read_df_from_bytes(file_bytes, catalog_file.name)
        st.session_state.system_file_name = catalog_file.name
        st.session_state.system_embedded_images = extract_embedded_images(file_bytes, catalog_file.name)

    if "system_df" in st.session_state:
        df_catalog = st.session_state.system_df
        embedded_map = st.session_state.get("system_embedded_images", {})
        embedded_cols = {c for (_, c) in embedded_map.keys()}

        auto_name_cols = guess_name_columns(df_catalog)
        auto_pack_cols = guess_image_columns(df_catalog, embedded_cols)

        st.success(
            f"🤖 Auto-detected — Naam/Group: **{', '.join(auto_name_cols) or 'nahi mila'}** · "
            f"Packshot: **{', '.join(auto_pack_cols) or 'nahi mila'}**"
        )
        if embedded_cols:
            st.caption(f"📎 Is file me embedded (seedhi chipki hui) images mili: {', '.join(embedded_cols)}")

        with st.expander("⚙️ Advanced: columns khud choose karein (zaroori nahi hai)"):
            system_name_cols = st.multiselect(
                "Naam/Group/Class column(s)", list(df_catalog.columns), default=auto_name_cols, key="sys_name_cols"
            )
            system_pack_cols = st.multiselect(
                "Packshot column(s)", list(df_catalog.columns), default=auto_pack_cols, key="sys_pack_cols"
            )

        resolved_name_cols = st.session_state.get("sys_name_cols", auto_name_cols)
        resolved_pack_cols = st.session_state.get("sys_pack_cols", auto_pack_cols)

        search = st.text_input("🔍 Search (naam se dhoondhein)")
        view_df = df_catalog.copy()
        if search:
            mask = view_df.apply(lambda r: search.lower() in " ".join(str(v) for v in r.values).lower(), axis=1)
            view_df = view_df[mask]

        preview_rows = []
        for idx, row in view_df.iterrows():
            pack = get_packshot_value(row, idx, resolved_pack_cols, embedded_map)
            preview_rows.append({**row.to_dict(), "Packshot Preview": pack["display"]})
        display_df = pd.DataFrame(preview_rows) if preview_rows else view_df

        column_config = {}
        if resolved_pack_cols:
            column_config["Packshot Preview"] = st.column_config.ImageColumn("Packshot", width="small")

        st.dataframe(display_df, use_container_width=True, hide_index=True, column_config=column_config)
        st.caption(f"Total {len(df_catalog)} entries system me hain — {len(view_df)} dikh rahi hain.")
    else:
        st.info("Upar se apni system sheet upload karke shuru karein.")

# =====================================================================
# TOOL 4: SKU Match   (fully automatic, embedded-image aware, any format)
# =====================================================================

else:
    st.header("SKU Match")
    st.caption(
        "Client ki sheet kaisi bhi ho, tool khud naam aur packshot column pehchan kar match karega — "
        "image ho link se ya sheet me seedhi chipki hui (embedded), dono chalengi."
    )

    if "system_df" not in st.session_state:
        st.warning("Pehle **'System SKU Catalog'** tab me apni system sheet upload kar lein, phir yahan match karein.")
    else:
        df_internal = st.session_state.system_df
        internal_embedded = st.session_state.get("system_embedded_images", {})
        internal_name_cols = st.session_state.get(
            "sys_name_cols", guess_name_columns(df_internal)
        )
        internal_pack_cols = st.session_state.get(
            "sys_pack_cols", guess_image_columns(df_internal, {c for (_, c) in internal_embedded.keys()})
        )

        st.info(
            f"System catalog: **{st.session_state.get('system_file_name','')}** "
            f"({len(df_internal)} rows) — 'System SKU Catalog' tab se aata hai."
        )

        if not internal_name_cols:
            st.warning("'System SKU Catalog' tab me naam wala column detect/select nahi ho paaya.")
        else:
            client_file = st.file_uploader(
                "Client ki SKU file (kisi bhi format me — tool khud samajh lega)",
                type=["xlsx", "xls", "csv"],
                key="client_file2",
            )

            if client_file:
                client_bytes = get_file_bytes(client_file)
                df_client = read_df_from_bytes(client_bytes, client_file.name)
                client_embedded = extract_embedded_images(client_bytes, client_file.name)
                client_embedded_cols = {c for (_, c) in client_embedded.keys()}

                auto_client_name = guess_name_columns(df_client)
                auto_client_pack = guess_image_columns(df_client, client_embedded_cols)

                st.success(
                    f"🤖 Auto-detected — Naam/Group: **{', '.join(auto_client_name) or 'nahi mila'}** · "
                    f"Packshot: **{', '.join(auto_client_pack) or 'nahi mila'}**"
                )
                if client_embedded_cols:
                    st.caption(f"📎 Is file me embedded images mili: {', '.join(client_embedded_cols)}")

                with st.expander("⚙️ Advanced: client ke columns khud choose karein (zaroori nahi hai)"):
                    client_name_cols = st.multiselect(
                        "Client Naam/Group column(s)", list(df_client.columns), default=auto_client_name, key="client_name_cols"
                    )
                    client_pack_cols = st.multiselect(
                        "Client Packshot column(s) (Front/Back — jitne bhi hon)",
                        list(df_client.columns), default=auto_client_pack, key="client_pack_cols",
                    )

                resolved_client_name = st.session_state.get("client_name_cols", auto_client_name)
                resolved_client_pack = st.session_state.get("client_pack_cols", auto_client_pack)

                threshold = st.slider(
                    "Fuzzy match sensitivity (naam thoda alag ho tab bhi match kare)", 50, 95, 70, step=5, format="%d%%"
                ) / 100.0
                embed_in_download = st.checkbox(
                    "Downloaded Excel me actual packshot images bhi chipka dein (jahan sheet me photo pehle se embedded ho)",
                    value=True,
                )

                if st.button("🔎 Match karein", type="primary", disabled=not resolved_client_name):
                    internal_records = []
                    for idx, irow in df_internal.iterrows():
                        combined = combine_cols(irow, internal_name_cols)
                        pack = get_packshot_value(irow, idx, internal_pack_cols, internal_embedded)
                        internal_records.append((combined, pack))

                    exact_map = {}
                    for combined, pack in internal_records:
                        norm = normalize(combined)
                        if norm not in exact_map:
                            exact_map[norm] = (combined, pack)

                    results = []
                    client_pack_bytes_map = {}
                    matched_pack_bytes_map = {}

                    for idx, crow in df_client.iterrows():
                        client_combined = combine_cols(crow, resolved_client_name)
                        norm = normalize(client_combined)
                        client_pack = get_packshot_value(crow, idx, resolved_client_pack, client_embedded)

                        match_type, matched_val, confidence = "Unmatched", "", 0.0
                        matched_pack = {"display": "", "bytes": None}

                        if norm in exact_map:
                            matched_val, matched_pack = exact_map[norm]
                            match_type, confidence = "Exact", 1.0
                        else:
                            best_score, best_combined, best_pack = 0.0, None, {"display": "", "bytes": None}
                            for combined, pack in internal_records:
                                sc = similarity(client_combined, combined)
                                if sc > best_score:
                                    best_score, best_combined, best_pack = sc, combined, pack
                            if best_combined is not None and best_score >= threshold:
                                matched_val, match_type, confidence, matched_pack = best_combined, "Fuzzy", best_score, best_pack

                        row_num = len(results)
                        if client_pack["bytes"]:
                            client_pack_bytes_map[row_num] = client_pack["bytes"]
                        if matched_pack["bytes"]:
                            matched_pack_bytes_map[row_num] = matched_pack["bytes"]

                        results.append({
                            "Client SKU/Group": client_combined,
                            "Client Packshot": client_pack["display"],
                            "Matched SKU (System)": matched_val,
                            "Matched Packshot": matched_pack["display"],
                            "Match Type": match_type,
                            "Confidence %": round(confidence * 100),
                        })

                    st.session_state.match_results = pd.DataFrame(results)
                    st.session_state.match_client_pack_bytes = client_pack_bytes_map
                    st.session_state.match_matched_pack_bytes = matched_pack_bytes_map
                    st.session_state.match_embed_flag = embed_in_download

                if "match_results" in st.session_state:
                    results_df = st.session_state.match_results
                    total = len(results_df)
                    matched_count = int((results_df["Match Type"] != "Unmatched").sum()) if total else 0
                    match_rate = round(matched_count / total * 100) if total else 0
                    counts = results_df["Match Type"].value_counts().to_dict()

                    m0, m1, m2, m3 = st.columns(4)
                    m0.metric("📊 Match Rate", f"{match_rate}%")
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

                    st.dataframe(
                        show_df,
                        use_container_width=True,
                        hide_index=True,
                        column_config={
                            "Client Packshot": st.column_config.ImageColumn("Client Packshot", width="small"),
                            "Matched Packshot": st.column_config.ImageColumn("Matched Packshot", width="small"),
                            "Confidence %": st.column_config.ProgressColumn(
                                "Confidence", min_value=0, max_value=100, format="%d%%"
                            ),
                        },
                    )

                    embed_flag = st.session_state.get("match_embed_flag", True)
                    excel_bytes = build_match_excel(
                        results_df,
                        {
                            "client": st.session_state.get("match_client_pack_bytes", {}) if embed_flag else {},
                            "matched": st.session_state.get("match_matched_pack_bytes", {}) if embed_flag else {},
                        },
                    )
                    st.download_button(
                        "⬇ Report download",
                        data=excel_bytes,
                        file_name="sku-match-report.xlsx",
                        mime=EXCEL_MIME,
                    )
