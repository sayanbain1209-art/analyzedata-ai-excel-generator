import io
import re
import csv
from datetime import date

import pandas as pd
import streamlit as st
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

st.set_page_config(page_title="AnalyzeData.io | AI Excel Generator", page_icon="📊", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
:root { --navy:#12345a; --blue:#2563eb; }
.stApp { background:#f6f8fc; color:#172b45; }
[data-testid="stHeader"] { background:rgba(246,248,252,.95); }
[data-testid="stSidebar"] { background:#102b4c; }
[data-testid="stSidebar"] * { color:#f8fbff !important; }
.stApp, .stApp p, .stApp label, .stApp h1, .stApp h2, .stApp h3, .stApp [data-testid="stMarkdownContainer"] { color:#172b45; }
.stApp input, .stApp textarea, .stApp [data-baseweb="select"] * { color:#172b45 !important; }
.hero { padding:1.4rem 1.5rem; border-radius:18px; background:linear-gradient(120deg,#102b4c,#2563eb); color:white; margin-bottom:1rem; }
.hero h1,.hero p { color:white !important; }
.metric-card { background:white; border:1px solid #e5eaf2; border-radius:14px; padding:1rem; min-height:95px; }
.metric-label { color:#64748b; font-size:.86rem; }.metric-value { color:#12345a; font-size:1.5rem; font-weight:750; }
.hint { color:#526783; font-size:.95rem; }
div.stButton > button[kind="primary"] { background:#2563eb; border-color:#2563eb; border-radius:10px; font-weight:650; min-height:2.8rem; }
</style>
""", unsafe_allow_html=True)

SALARY_COLUMNS = ["Employee ID", "Employee Name", "Department", "Basic Salary", "HRA", "Other Allowance", "Gross Salary", "Deduction", "Net Salary"]


def prompt_to_columns(prompt):
    match = re.search(r"(?:columns?|fields?)\s*[:\-]\s*(.+)", prompt, flags=re.I | re.S)
    if not match:
        return []
    candidate = re.sub(r"[.。].*$", "", match.group(1).split("\n")[0])
    cols = [x.strip(" \t:-•") for x in re.split(r",|;|\band\b|&", candidate, flags=re.I)]
    return list(dict.fromkeys([x for x in cols if 1 < len(x) < 50]))[:25]


def parse_pasted_data(raw):
    """Read pasted CSV, tab-separated, or pipe-separated rows into a dataframe."""
    raw = raw.strip()
    if not raw:
        raise ValueError("প্রথমে কর্মচারীদের data paste করুন।")
    lines = [line for line in raw.splitlines() if line.strip()]
    sample = "\n".join(lines[:5])
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters="\t,;|")
        delimiter = dialect.delimiter
    except csv.Error:
        delimiter = "\t" if "\t" in sample else ("|" if "|" in sample else ",")
    rows = list(csv.reader(io.StringIO("\n".join(lines)), delimiter=delimiter))
    rows = [[cell.strip() for cell in row] for row in rows]
    rows = [row for row in rows if any(row)]
    if len(rows) < 2:
        raise ValueError("Header-সহ অন্তত ১টি data row দিন। প্রতিটি column tab, comma বা | দিয়ে আলাদা করুন।")
    width = len(rows[0])
    rows = [row for row in rows if len(row) == width]
    if len(rows) < 2:
        raise ValueError("সব row-তে একই সংখ্যক column নেই। আবার paste করুন।")
    headers = rows[0]
    df = pd.DataFrame(rows[1:], columns=headers)
    return clean_and_calculate(df)


def clean_and_calculate(df):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    # Normalize common salary column names.
    aliases = {
        "name": "Employee Name", "employee name": "Employee Name", "employee id": "Employee ID",
        "basic": "Basic Salary", "basic salary": "Basic Salary", "hra": "HRA",
        "other allowance": "Other Allowance", "allowance": "Other Allowance",
        "deduction": "Deduction", "department": "Department"
    }
    rename = {c: aliases[c.strip().lower()] for c in df.columns if c.strip().lower() in aliases}
    df = df.rename(columns=rename)
    for col in ["Basic Salary", "HRA", "Other Allowance", "Deduction"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col].astype(str).str.replace(",", "", regex=False).str.replace("₹", "", regex=False).str.strip(), errors="coerce").fillna(0)
    if all(c in df.columns for c in ["Basic Salary", "HRA", "Other Allowance", "Deduction"]):
        df["Gross Salary"] = df["Basic Salary"] + df["HRA"] + df["Other Allowance"]
        df["Net Salary"] = df["Gross Salary"] - df["Deduction"]
        ordered = [c for c in SALARY_COLUMNS if c in df.columns]
        extras = [c for c in df.columns if c not in ordered]
        df = df[ordered + extras]
    return df


def make_starter_dataframe(prompt, template):
    templates = {
        "Employee Details": ["Employee ID", "Employee Name", "Department", "Designation", "Phone", "Joining Date"],
        "Salary Calculator": SALARY_COLUMNS,
        "Attendance Sheet": ["Employee ID", "Employee Name", "Working Days", "Present Days", "Absent", "Attendance %"],
        "Student Marks": ["Student ID", "Student Name", "Subject", "Total Marks", "Obtained Marks", "Percentage"],
    }
    if template in templates:
        return pd.DataFrame(columns=templates[template])
    cols = prompt_to_columns(prompt)
    return pd.DataFrame(columns=cols or ["ID", "Name", "Category", "Date", "Notes"])


def parse_uploaded_file(uploaded_file):
    name, raw = uploaded_file.name.lower(), uploaded_file.getvalue()
    if name.endswith(".csv"):
        return clean_and_calculate(pd.read_csv(io.BytesIO(raw)))
    if name.endswith((".xlsx", ".xls")):
        return clean_and_calculate(pd.read_excel(io.BytesIO(raw)))
    raise ValueError("এই সংস্করণে CSV এবং XLSX ফাইল সমর্থিত।")


def style_workbook(df):
    output = io.BytesIO()
    safe_df = df.copy().fillna("")
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        safe_df.to_excel(writer, index=False, sheet_name="Salary Report" if "Net Salary" in safe_df.columns else "Generated Data")
        ws = writer.sheets[writer.book.sheetnames[0]]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        ws.sheet_view.showGridLines = False
        ws.row_dimensions[1].height = 30
        header_fill = PatternFill("solid", fgColor="17365D")
        header_font = Font(name="Aptos", bold=True, color="FFFFFF", size=11)
        edge = Side(style="thin", color="D7E2F0")
        for cell in ws[1]:
            cell.fill, cell.font = header_fill, header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = Border(left=edge, right=edge, top=edge, bottom=edge)
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.font = Font(name="Aptos", size=10, color="243247")
                cell.alignment = Alignment(vertical="center", wrap_text=True)
                cell.border = Border(left=edge, right=edge, top=edge, bottom=edge)
                if cell.row % 2 == 0:
                    cell.fill = PatternFill("solid", fgColor="F1F6FC")
        for idx, cells in enumerate(ws.columns, 1):
            longest = max([len(str(c.value or "")) for c in cells] + [10])
            ws.column_dimensions[get_column_letter(idx)].width = min(max(longest + 3, 14), 32)
        if ws.max_row >= 2 and ws.max_column >= 1:
            table = Table(displayName="GeneratedTable", ref=f"A1:{get_column_letter(ws.max_column)}{ws.max_row}")
            table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True, showColumnStripes=False)
            ws.add_table(table)
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth, ws.page_setup.fitToHeight, ws.page_setup.orientation = 1, 0, "landscape"
        # Add a department summary on a second sheet for salary reports.
        if all(c in safe_df.columns for c in ["Department", "Net Salary"]):
            summary = safe_df.groupby("Department", dropna=False)["Net Salary"].sum().reset_index(name="Total Net Salary")
            summary.to_excel(writer, index=False, sheet_name="Department Summary")
            sw = writer.sheets["Department Summary"]
            sw.freeze_panes = "A2"
            for cell in sw[1]:
                cell.fill, cell.font = header_fill, header_font
            for col_idx, cells in enumerate(sw.columns, 1):
                sw.column_dimensions[get_column_letter(col_idx)].width = max(18, max(len(str(c.value or "")) for c in cells) + 3)
    output.seek(0)
    return output.getvalue()


if "dataframe" not in st.session_state:
    st.session_state.dataframe = pd.DataFrame(columns=["ID", "Name", "Date", "Amount", "Notes"])
if "file_name" not in st.session_state:
    st.session_state.file_name = "AI_Generated_Excel.xlsx"

with st.sidebar:
    st.markdown("## 📊 AnalyzeData.io")
    st.caption("AI Excel Generator")
    page = st.radio("NAVIGATION", ["Dashboard", "AI Excel Generator", "Templates", "About"])
    st.divider()
    st.caption("CURRENT FEATURES")
    st.markdown("**Paste data:** Available")
    st.markdown("**Salary calculations:** Available")
    st.markdown("**CSV / XLSX import:** Available")
    st.markdown("**Gemini AI / PDF / Image:** Not connected yet")

if page == "Dashboard":
    st.markdown('<div class="hero"><h1>Turn your data into professional Excel files.</h1><p>Paste data, upload a file, edit the preview, and download a formatted workbook.</p></div>', unsafe_allow_html=True)
    a,b,c = st.columns(3)
    for col, label, value in [(a,"Workbook builder","Ready"),(b,"Editable preview","Included"),(c,"Salary calculations","Included")]:
        with col: st.markdown(f'<div class="metric-card"><div class="metric-label">{label}</div><div class="metric-value">{value}</div></div>', unsafe_allow_html=True)
    st.subheader("Start creating")
    st.write("বাম পাশের Navigation থেকে **AI Excel Generator** খুলুন।")

elif page == "AI Excel Generator":
    st.title("AI Excel Generator")
    st.caption("Paste your source data or upload a CSV/XLSX file. Edit the result and download Excel.")
    st.subheader("1. Paste your data")
    pasted = st.text_area("Paste table data here (header row first)", height=210, placeholder="Employee ID\tEmployee Name\tDepartment\tBasic Salary\tHRA\tOther Allowance\tDeduction\nEMP101\tRahul Das\tSales\t18000\t3000\t2000\t1000")
    if st.button("Convert pasted data to Excel table", type="primary", use_container_width=True):
        try:
            st.session_state.dataframe = parse_pasted_data(pasted)
            st.session_state.file_name = "Pasted_Data_Excel.xlsx"
            st.success("তোমার paste করা data টেবিলে চলে এসেছে। নিচে preview দেখো।")
        except Exception as exc:
            st.error(str(exc))
    st.divider()
    st.subheader("2. Or create an empty structure / upload a file")
    prompt = st.text_area("Describe the columns you want", placeholder="Create a sheet with columns: Name, Department, Phone, Joining Date", height=90)
    left,right = st.columns(2)
    with left:
        template = st.selectbox("Start with a template", ["Custom / Universal", "Employee Details", "Salary Calculator", "Attendance Sheet", "Student Marks"])
        if st.button("Generate empty structure", use_container_width=True):
            st.session_state.dataframe = make_starter_dataframe(prompt, template)
            st.session_state.file_name = "AI_Generated_Excel.xlsx"
            st.success("ফাঁকা structure তৈরি হয়েছে।")
    with right:
        upload = st.file_uploader("Upload source data (CSV / XLSX)", type=["csv", "xlsx", "xls"])
        if upload is not None and st.button("Load uploaded file", use_container_width=True):
            try:
                st.session_state.dataframe = parse_uploaded_file(upload)
                st.session_state.file_name = re.sub(r"\.[^.]+$", "", upload.name) + "_formatted.xlsx"
                st.success("ফাইল load হয়েছে।")
            except Exception as exc: st.error(str(exc))
    st.divider()
    st.subheader("3. Excel Preview & Editor")
    st.caption("টেবিলের cell-এ click করে data পরিবর্তন করতে বা নতুন row যোগ করতে পারো।")
    edited = st.data_editor(st.session_state.dataframe, num_rows="dynamic", use_container_width=True, hide_index=True, key="excel_editor_v2")
    edited = clean_and_calculate(edited)
    st.session_state.dataframe = edited
    if "Net Salary" in edited.columns and len(edited):
        st.metric("Total Net Salary", f"₹{pd.to_numeric(edited['Net Salary'], errors='coerce').fillna(0).sum():,.2f}")
        if "Department" in edited.columns:
            st.subheader("Department-wise Net Salary")
            st.dataframe(edited.groupby("Department", dropna=False)["Net Salary"].sum().reset_index(name="Total Net Salary"), use_container_width=True, hide_index=True)
    st.download_button("⬇️ Download Professional Excel", data=style_workbook(edited), file_name=st.session_state.file_name, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary", use_container_width=True)
    st.info("নোট: এই সংস্করণে pasted CSV/tab/pipe data, salary হিসাব, department summary এবং Excel export কাজ করে। Gemini-ভিত্তিক প্রকৃত AI, PDF পড়া ও image OCR এখনো যুক্ত করা হয়নি।")

elif page == "Templates":
    st.title("Ready-made Excel Templates")
    st.write("Generator পেজে template বেছে empty structure তৈরি করতে পারো।")
    for name, cols in {"Employee Details":"Employee ID, Employee Name, Department, Designation, Phone, Joining Date","Salary Calculator":", ".join(SALARY_COLUMNS),"Attendance Sheet":"Employee ID, Employee Name, Working Days, Present Days, Absent, Attendance %","Student Marks":"Student ID, Student Name, Subject, Total Marks, Obtained Marks, Percentage"}.items():
        with st.expander(name): st.write(cols)
else:
    st.title("About AnalyzeData.io")
    st.write("AnalyzeData.io text table বা CSV/XLSX data থেকে editable ও formatted Excel workbook তৈরি করে।")
    st.warning("Gemini AI, PDF parsing এবং image OCR এখনো সংযুক্ত নয়।")