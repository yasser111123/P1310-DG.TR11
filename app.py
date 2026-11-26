import io
import os
import re
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import streamlit as st

st.set_page_config(page_title="P1310 PDB DG.TR11", page_icon="🇱", layout="centered")

if "light_mode" not in st.session_state:
    st.session_state.light_mode = False

# Top bar sghira fiha Moon / Sun toggle button 3la limn
_, top_btn_col = st.columns([6, 1.4])
with top_btn_col:
    toggle_label = "🌙 Dark" if st.session_state.light_mode else "☀️ Light"
    if st.button(toggle_label, key="theme_toggle_btn"):
        st.session_state.light_mode = not st.session_state.light_mode
        st.rerun()

if st.session_state.light_mode:
    bg_color = "#ffffff"
    text_color = "#0f172a"
    heading_color = "#0f172a"
    textarea_bg = "#f8fafc"
    textarea_text = "#0f172a"
    textarea_border = "#cbd5e1"
    placeholder_color = "#64748b"
    select_bg = "#f8fafc"
    select_text = "#0f172a"
    select_border = "#cbd5e1"
    toggle_bg = "#f1f5f9"
    toggle_text = "#0f172a"
    toggle_border = "#cbd5e1"
    logo_color = "#13294b"
    sub_color = "#475569"
    th_bg = "#e2e8f0"
    th_color = "#13294b"
    th_border = "#cbd5e1"
    td_bg = "#ffffff"
    td_color = "#1e293b"
    td_border = "#e2e8f0"
    hr_color = "#e2e8f0"
else:
    bg_color = "#0b0f19"
    text_color = "#e2e8f0"
    heading_color = "#e2e8f0"
    textarea_bg = "#111827"
    textarea_text = "#f3f4f6"
    textarea_border = "#1f2937"
    placeholder_color = "#94a3b8"
    select_bg = "#111827"
    select_text = "#f3f4f6"
    select_border = "#1f2937"
    toggle_bg = "#111827"
    toggle_text = "#f3f4f6"
    toggle_border = "#1f2937"
    logo_color = "#ffffff"
    sub_color = "#94a3b8"
    th_bg = "#1e293b"
    th_color = "#38bdf8"
    th_border = "#334155"
    td_bg = "#0f172a"
    td_color = "#cbd5e1"
    td_border = "#1e293b"
    hr_color = "#1f2937"

st.markdown(f"""
    <style>
    .stApp, .main, [data-testid="stAppViewContainer"] {{
        background-color: {bg_color} !important;
        color: {text_color} !important;
    }}

    h1, h2, h3, h4, h5, h6, p, label, .stMarkdown {{
        color: {heading_color} !important;
    }}

    hr {{
        border-color: {hr_color} !important;
    }}

    .stTextArea textarea {{
        background-color: {textarea_bg} !important;
        color: {textarea_text} !important;
        border: 1px solid {textarea_border} !important;
        font-family: 'Consolas', monospace !important;
        font-size: 15px !important;
    }}
    .stTextArea textarea::placeholder {{
        color: {placeholder_color} !important;
        opacity: 1 !important;
    }}
    .stTextArea textarea:focus {{
        border-color: #3b82f6 !important;
        box-shadow: 0 0 8px rgba(59, 130, 246, 0.4) !important;
    }}

    div[data-baseweb="select"] > div {{
        background-color: {select_bg} !important;
        color: {select_text} !important;
        border: 1px solid {select_border} !important;
        border-radius: 8px !important;
    }}

    div.stButton > button[kind="secondary"],
    div.stButton > button[kind="primary"],
    div.stButton > button:first-child,
    div[data-testid="stDownloadButton"] > button {{
        background-color: #2563eb !important;
        color: #ffffff !important;
        font-weight: bold !important;
        width: 100%;
        border: none !important;
        padding: 12px !important;
        border-radius: 8px !important;
        transition: all 0.3s ease !important;
    }}
    div.stButton > button:first-child:hover,
    div[data-testid="stDownloadButton"] > button:hover {{
        background-color: #1d4ed8 !important;
        box-shadow: 0 0 12px rgba(37, 99, 235, 0.5) !important;
    }}

    .st-key-theme_toggle_btn button {{
        background-color: {toggle_bg} !important;
        color: {toggle_text} !important;
        border: 1px solid {toggle_border} !important;
        padding: 6px 10px !important;
        font-size: 13px !important;
        border-radius: 999px !important;
    }}
    .st-key-theme_toggle_btn button:hover {{
        border-color: #3b82f6 !important;
        box-shadow: none !important;
    }}

    .brand-logo-container {{
        text-align: center;
        margin-bottom: 25px;
    }}
    .brand-logo {{
        font-family: 'Arial Black', sans-serif;
        font-size: 54px;
        font-weight: 900;
        color: {logo_color} !important;
        letter-spacing: 4px;
        margin-bottom: -5px;
        transition: color 0.25s ease;
    }}
    .brand-sub {{
        color: {sub_color} !important;
        font-family: 'Consolas', monospace;
        font-size: 13px;
        font-weight: bold;
        letter-spacing: 1.5px;
    }}

    th {{
        background-color: {th_bg} !important;
        color: {th_color} !important;
        text-align: center !important;
        border-bottom: 2px solid {th_border} !important;
    }}
    td {{
        background-color: {td_bg} !important;
        color: {td_color} !important;
        text-align: center !important;
        border-bottom: 1px solid {td_border} !important;
    }}
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


def build_export_xlsx(details_rows, summary_rows):
    wb = openpyxl.Workbook()
    header_fill = PatternFill(start_color="13294B", end_color="13294B", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    total_fill = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
    total_font = Font(name="Calibri", size=11, bold=True, color="0F172A")
    body_font = Font(name="Calibri", size=11, color="0F172A")
    thin_side = Side(border_style="thin", color="CBD5E1")
    cell_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    def write_sheet(ws, rows_data):
        if not rows_data:
            ws.append(["No Data"])
            return

        headers = list(rows_data[0].keys())
        ws.append(headers)

        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(1, col_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = cell_border
        ws.row_dimensions[1].height = 24

        for r_idx, row_dict in enumerate(rows_data, start=2):
            vals = [row_dict.get(h, "") for h in headers]
            ws.append(vals)
            is_total = str(vals[0]).strip().upper() == "TOTAL"
            ws.row_dimensions[r_idx].height = 20

            for c_idx in range(1, len(headers) + 1):
                c_cell = ws.cell(r_idx, c_idx)
                c_cell.font = total_font if is_total else body_font
                if is_total:
                    c_cell.fill = total_fill
                c_cell.alignment = Alignment(horizontal="center", vertical="center")
                c_cell.border = cell_border

        for col_cells in ws.columns:
            col_letter = openpyxl.utils.get_column_letter(col_cells[0].column)
            max_len = max(len(str(cell.value or "")) for cell in col_cells)
            ws.column_dimensions[col_letter].width = max(max_len + 6, 14)

    ws_details = wb.active
    ws_details.title = "References Details"
    write_sheet(ws_details, details_rows)

    ws_summary = wb.create_sheet(title="Versions And Quantity")
    write_sheet(ws_summary, summary_rows)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


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
        "Filter EP",
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
                        "EP": "-",
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
                        "EP": "-",
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
                            "EP": cond_name,
                            "Version": ver_name,
                            "Qty": qty,
                        })

            if not found_any_row and not selected_conds:
                matches_table.append({
                    "Input": search_d,
                    "Reference": matched_cols[0]["full_ref"],
                    "EP": "Makayn f hta EP",
                    "Version": "-",
                    "Qty": qty,
                })

        st.markdown("---")
        st.markdown("### 🎯 References Details:")
        display_details = matches_table if matches_table else [{"Input": "-", "Reference": "Null", "EP": "-", "Version": "-", "Qty": 0}]
        st.table(display_details)

        st.markdown("### 📊Versions And Quantity ")
        summary_rows = []
        for (cond_name, ver_name), q_sum in sorted(
            cond_ver_qty.items(),
            key=lambda item: (item[0][0], int(item[0][1]) if item[0][1].isdigit() else 9999)
        ):
            wires_lr = data["wires_map"].get((cond_name, ver_name), "-")
            summary_rows.append({
                "EP": cond_name,
                "Version": ver_name,
                "Total Qty": q_sum,
                "Wires (L/R)": wires_lr,
            })

        if summary_rows:
            summary_rows.append({
                "EP": "TOTAL",
                "Version": "-",
                "Total Qty": total_qty,
                "Wires (L/R)": "-",
            })
        else:
            summary_rows.append({
                "EP": "Null",
                "Version": "-",
                "Total Qty": 0,
                "Wires (L/R)": "-",
            })

        st.table(summary_rows)

        xlsx_bytes = build_export_xlsx(display_details, summary_rows)
        st.download_button(
            label="📥 Export Excel (.xlsx)",
            data=xlsx_bytes,
            file_name="DG_TR11_Synthese_Export.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
