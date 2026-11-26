import os
import re
import openpyxl
import streamlit as st

st.set_page_config(page_title="P1310 PDB DG.TR11", page_icon="🇱", layout="centered")

st.markdown("""
    <style>
    .main { background-color: #0b0f19; color: #e2e8f0; }

    .stTextArea textarea {
        background-color: #111827 !important;
        color: #f3f4f6 !important;
        border: 1px solid #1f2937 !important;
        font-family: 'Consolas', monospace !important;
    }
    .stTextArea textarea:focus {
        border-color: #3b82f6 !important;
        box-shadow: 0 0 8px rgba(59, 130, 246, 0.4) !important;
    }

    div[data-baseweb="select"] > div {
        background-color: #111827 !important;
        color: #f3f4f6 !important;
        border: 1px solid #1f2937 !important;
        border-radius: 8px !important;
    }

    div.stButton > button:first-child {
        background-color: #2563eb !important;
        color: #ffffff !important;
        font-weight: bold !important;
        width: 100%;
        border: none !important;
        padding: 12px !important;
        border-radius: 8px !important;
        transition: all 0.3s ease !important;
    }
    div.stButton > button:first-child:hover {
        background-color: #1d4ed8 !important;
        box-shadow: 0 0 12px rgba(37, 99, 235, 0.5) !important;
    }

    .brand-logo-container {
        text-align: center;
        margin-bottom: 25px;
    }
    .brand-logo {
        font-family: 'Arial Black', sans-serif;
        font-size: 54px;
        font-weight: 900;
        color: #ffffff;
        letter-spacing: 4px;
        margin-bottom: -5px;
    }
    .brand-logo span {
        color: #38bdf8; /* T-touch dyal l'blue ila bghiti */
    }
    .brand-sub {
        color: #94a3b8;
        font-family: 'Consolas', monospace;
        font-size: 13px;
        font-weight: bold;
        letter-spacing: 1.5px;
    }

    th {
        background-color: #1e293b !important;
        color: #38bdf8 !important;
        text-align: center !important;
        border-bottom: 2px solid #334155 !important;
    }
    td {
        background-color: #0f172a !important;
        color: #cbd5e1 !important;
        text-align: center !important;
        border-bottom: 1px solid #1e293b !important;
    }
    </style>
""", unsafe_allow_html=True)

EXCEL_FILE = "DG.TR11.  Splices_X1310_HARN DASHBOARD DG(1).xlsx"
SHEET_NAME = "Synthese"

def clean_digits(val):
    if val is None:
        return ""
    s = str(val).strip()
    if len(s) > 4 and s[-1].isalpha() and "--" not in s:
        s = s[:-1]
    return "".join(c for c in s if c.isdigit())

def format_num_val(val):
    if val is None or str(val).strip() == "":
        return "-"
    if isinstance(val, (int, float)) and float(val).is_integer():
        return str(int(val))
    return str(val).strip()

def extract_ref_and_qty(line_text):
    tokens = re.findall(r"\d+[a-zA-Z0-9_-]*", line_text)
    if not tokens:
        return 1, ""

    qty_sum = 0
    ref_token = ""

    for tok in tokens:
        digits = "".join(c for c in tok if c.isdigit())
        if len(digits) >= 4:
            ref_token = tok
        elif digits.isdigit():
            qty_sum += int(digits)

    if not ref_token:
        ref_token = tokens[-1]
        digits_last = "".join(c for c in ref_token if c.isdigit())
        if digits_last.isdigit() and len(digits_last) < 4 and qty_sum >= int(digits_last):
            qty_sum -= int(digits_last)

    if qty_sum == 0:
        qty_sum = 1

    return qty_sum, ref_token

@st.cache_data
def load_synthese_data(excel_path):
    if not os.path.exists(excel_path):
        return None

    wb = openpyxl.load_workbook(excel_path, data_only=True)
    if SHEET_NAME not in wb.sheetnames:
        return None

    ws = wb[SHEET_NAME]

    col_map = {}
    for c in range(1, ws.max_column + 1):
        header_val = ws.cell(2, c).value
        if header_val is not None:
            col_map[str(header_val).strip()] = c

    cond_col = col_map.get("Cond", 3)
    version_col = col_map.get("Version", 4)
    wires_left_col = col_map.get("Wires Left", 10)
    wires_right_col = col_map.get("Wires Right", 11)
    n_ref_col = col_map.get("N ref(s)", 18)

    ref_columns = []
    for c in range(n_ref_col + 1, ws.max_column + 1):
        full_ref = ws.cell(1, c).value
        short_ref = ws.cell(2, c).value
        if full_ref is not None or short_ref is not None:
            full_str = str(full_ref).strip() if full_ref is not None else str(short_ref).strip()
            ref_columns.append({
                "col": c,
                "full_ref": full_str,
                "short_ref": str(short_ref).strip() if short_ref is not None else "",
                "digits": clean_digits(full_str if full_ref is not None else short_ref),
            })

    rows = []
    wires_map = {}
    cond_options = set()

    for r in range(3, ws.max_row + 1):
        cond_val = ws.cell(r, cond_col).value
        if cond_val is None or str(cond_val).strip() == "":
            continue

        cond_str = str(cond_val).strip()
        cond_options.add(cond_str)
        ver_str = format_num_val(ws.cell(r, version_col).value)
        w_left = format_num_val(ws.cell(r, wires_left_col).value)
        w_right = format_num_val(ws.cell(r, wires_right_col).value)
        wires_map[(cond_str, ver_str)] = f"{w_left}/{w_right}"

        active_refs = []
        for rc in ref_columns:
            cell_val = ws.cell(r, rc["col"]).value
            if cell_val is not None and str(cell_val).strip() != "" and str(cell_val).strip() != "0":
                active_refs.append(rc)

        rows.append({
            "cond": cond_str,
            "version": ver_str,
            "active_refs": active_refs,
        })

    return {
        "ref_columns": ref_columns,
        "rows": rows,
        "wires_map": wires_map,
        "cond_options": sorted(cond_options),
    }

# L-header dyal LEONI mzyan f-blasset l-9dima
st.markdown('<div class="brand-logo-container"><div class="brand-logo">LEONI</div><div class="brand-sub">P1310 PDB DG.TR11 </div></div>', unsafe_allow_html=True)

data = load_synthese_data(EXCEL_FILE)
cond_choices = data["cond_options"] if data else []

st.markdown("### 📝 Write down the reference with quantity :")
input_text = st.text_area(
    "",
    height=160,
    placeholder="Example :\n0122 126\n0510 21"
)

col_btn, col_filter = st.columns([1.2, 1])

with col_btn:
    run_clicked = st.button("Press to Continue ")

with col_filter:
    selected_conds = st.multiselect(
        "Filter E (Cond)",
        options=cond_choices,
        placeholder="Select The EPS",
        label_visibility="collapsed"
    )

if run_clicked:
    if data is None:
        st.error(f"❌ L-fichier '{EXCEL_FILE}' ma l9inahch f nfs l-dossier dyal `app.py`.")
    elif not input_text.strip():
        st.warning("⚠️ Kteb chi reference bhal `5117` ba3da.")
    else:
        lines = [ln.strip() for ln in input_text.strip().splitlines() if ln.strip()]
        matches_table = []
        cond_ver_qty = {}
        total_qty = 0

        for line in lines:
            qty, raw_ref = extract_ref_and_qty(line)
            search_d = clean_digits(raw_ref)

            if not search_d:
                if not selected_conds:
                    matches_table.append({
                        "Input": line,
                        "Reference": "❌ Ghalt",
                        "Cond": "-",
                        "Version": "-",
                        "Qty": qty,
                    })
                continue

            matched_cols = [
                rc for rc in data["ref_columns"]
                if rc["digits"].endswith(search_d) or search_d.endswith(rc["digits"])
            ]

            if not matched_cols:
                if not selected_conds:
                    matches_table.append({
                        "Input": search_d,
                        "Reference": f"❌ {search_d} ma kaynch",
                        "Cond": "-",
                        "Version": "-",
                        "Qty": qty,
                    })
                continue

            matched_col_indices = {mc["col"] for mc in matched_cols}
            found_any_row = False

            for row in data["rows"]:
                cond_name = row["cond"]
                if selected_conds and cond_name not in selected_conds:
                    continue

                for active in row["active_refs"]:
                    if active["col"] in matched_col_indices:
                        found_any_row = True
                        ver_name = row["version"]
                        key = (cond_name, ver_name)
                        cond_ver_qty[key] = cond_ver_qty.get(key, 0) + qty
                        total_qty += qty
                        matches_table.append({
                            "Input": search_d,
                            "Reference": active["full_ref"],
                            "Cond": cond_name,
                            "Version": ver_name,
                            "Qty": qty,
                        })

            if not found_any_row and not selected_conds:
                matches_table.append({
                    "Input": search_d,
                    "Reference": matched_cols[0]["full_ref"],
                    "Cond": "Makayn f hta Cond",
                    "Version": "-",
                    "Qty": qty,
                })

        st.markdown("---")
        st.markdown("### 🎯 References Details:")
        st.table(matches_table if matches_table else [{"Input": "-", "Reference": "Null", "Cond": "-", "Version": "-", "Qty": 0}])

        st.markdown("### 📊Versions And Quantity ")
        summary_rows = []
        for (cond_name, ver_name), q_sum in sorted(
            cond_ver_qty.items(),
            key=lambda item: (item[0][0], int(item[0][1]) if item[0][1].isdigit() else 9999)
        ):
            wires_lr = data["wires_map"].get((cond_name, ver_name), "-")
            summary_rows.append({
                "Cond": cond_name,
                "Version": ver_name,
                "Total Qty": q_sum,
                "Wires (L/R)": wires_lr,
            })

        if summary_rows:
            summary_rows.append({
                "Cond": "TOTAL",
                "Version": "-",
                "Total Qty": total_qty,
                "Wires (L/R)": "-",
            })
        else:
            summary_rows.append({
                "Cond": "Null",
                "Version": "-",
                "Total Qty": 0,
                "Wires (L/R)": "-",
            })

        st.table(summary_rows)
