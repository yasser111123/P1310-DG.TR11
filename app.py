import os
import re
import openpyxl
import streamlit as st

st.set_page_config(page_title="DG.TR11 - E01/E02/E03 PROG", page_icon="🚀", layout="centered")

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

    .brand-title {
        color: #ffffff;
        font-family: 'Arial Black', sans-serif;
        font-size: 28px;
        font-weight: 900;
        text-align: center;
        letter-spacing: 2px;
        margin-bottom: 0px;
    }
    .brand-sub {
        color: #94a3b8;
        font-family: 'Consolas', monospace;
        font-size: 13px;
        font-weight: bold;
        text-align: center;
        margin-bottom: 25px;
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
    splice_col = col_map.get("Splice", 2)
    version_col = col_map.get("Version", 4)
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
    all_conds = {}

    for r in range(3, ws.max_row + 1):
        cond_val = ws.cell(r, cond_col).value
        if cond_val is None or str(cond_val).strip() == "":
            continue

        cond_str = str(cond_val).strip()
        splice_str = str(ws.cell(r, splice_col).value or "").strip()
        ver_raw = ws.cell(r, version_col).value
        ver_str = (
            str(int(ver_raw))
            if isinstance(ver_raw, (int, float)) and float(ver_raw).is_integer()
            else str(ver_raw or "").strip()
        )

        all_conds.setdefault(cond_str, set())
        if ver_str:
            all_conds[cond_str].add(ver_str)

        active_refs = []
        for rc in ref_columns:
            cell_val = ws.cell(r, rc["col"]).value
            if cell_val is not None and str(cell_val).strip() != "" and str(cell_val).strip() != "0":
                active_refs.append(rc)

        rows.append({
            "cond": cond_str,
            "splice": splice_str,
            "version": ver_str,
            "active_refs": active_refs,
        })

    return {
        "ref_columns": ref_columns,
        "rows": rows,
        "all_conds": all_conds,
    }


st.markdown('<div class="brand-title">DG.TR11 - SYNTHESE</div>', unsafe_allow_html=True)
st.markdown('<div class="brand-sub">E01 • E02 • E03 ... VERSIONS & QUANTITY</div>', unsafe_allow_html=True)

st.markdown("### 📝 Kteb Reference o Quantity hna:")
input_text = st.text_area(
    "",
    height=160,
    placeholder="Mital:\n5117\n21 147 5117\n10 4733"
)

if st.button("🚀 Tla3 E01 / Versions o Hseb Quantity"):
    data = load_synthese_data(EXCEL_FILE)

    if data is None:
        st.error(f"❌ L-fichier '{EXCEL_FILE}' ma l9inahch f nfs l-dossier dyal `app.py`.")
    elif not input_text.strip():
        st.warning("⚠️ Kteb chi reference bhal `5117` ba3da.")
    else:
        lines = [ln.strip() for ln in input_text.strip().splitlines() if ln.strip()]
        matches_table = []
        cond_qty_totals = {}
        total_matched_qty = 0

        for line in lines:
            qty, raw_ref = extract_ref_and_qty(line)
            search_d = clean_digits(raw_ref)

            if not search_d:
                matches_table.append({
                    "Input": line,
                    "Reference": "❌ Ghalt",
                    "Cond": "-",
                    "Splice": "-",
                    "Version": "-",
                    "Qty": qty,
                })
                continue

            matched_cols = [
                rc for rc in data["ref_columns"]
                if rc["digits"].endswith(search_d) or search_d.endswith(rc["digits"])
            ]

            if not matched_cols:
                matches_table.append({
                    "Input": search_d,
                    "Reference": f"❌ {search_d} ma kaynch",
                    "Cond": "-",
                    "Splice": "-",
                    "Version": "-",
                    "Qty": qty,
                })
                continue

            matched_col_indices = {mc["col"] for mc in matched_cols}
            found_any_row = False

            for row in data["rows"]:
                for active in row["active_refs"]:
                    if active["col"] in matched_col_indices:
                        found_any_row = True
                        cond_name = row["cond"]
                        cond_qty_totals[cond_name] = cond_qty_totals.get(cond_name, 0) + qty
                        total_matched_qty += qty
                        matches_table.append({
                            "Input": search_d,
                            "Reference": active["full_ref"],
                            "Cond": cond_name,
                            "Splice": row["splice"],
                            "Version": row["version"],
                            "Qty": qty,
                        })

            if not found_any_row:
                matches_table.append({
                    "Input": search_d,
                    "Reference": matched_cols[0]["full_ref"],
                    "Cond": "Makayn f hta Cond",
                    "Splice": "-",
                    "Version": "-",
                    "Qty": qty,
                })

        st.markdown("---")
        st.markdown("### 🎯 Les `Cond` (E01, E02...) o `Version` li fihom had Reference:")
        st.table(matches_table)

        st.markdown("### 📊 Hasila dyal Quantity b-kol `Cond` (`E01`, `E02`, `E03`...):")
        all_e_summary = []
        for cond_name in sorted(data["all_conds"].keys()):
            vers_sorted = sorted(
                data["all_conds"][cond_name],
                key=lambda x: int(x) if x.isdigit() else 9999
            )
            qty_for_cond = cond_qty_totals.get(cond_name, 0)
            all_e_summary.append({
                "Cond": cond_name,
                "Versions": ", ".join(vers_sorted) if vers_sorted else "-",
                "Total Qty": qty_for_cond,
            })

        all_e_summary.append({
            "Cond": "TOTAL",
            "Versions": "-",
            "Total Qty": total_matched_qty,
        })

        st.table(all_e_summary)