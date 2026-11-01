import io
import re
from datetime import date

import pandas as pd
import streamlit as st
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

st.set_page_config(
    page_title="AnalyzeData.io | AI Excel Generator",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ---------- Theme ----------
st.markdown(
    """
    <style>
    :root { --navy:#12345a; --blue:#2563eb; --pale:#eff6ff; }
    .stApp { background: #f6f8fc; }
    [data-testid="stSidebar"] { background: #102b4c; }
    [data-testid="stSidebar"] * { color: #f8fbff !important; }
    .brand { font-size: 1.45rem; font-weight: 800; color:#12345a; letter-spacing:-.5px; }
    .hero {
        padding: 1.35rem 1.5rem; border-radius: 18px;
        background: linear-gradient(120deg,#102b4c,#2563eb);
        color: white; margin-bottom: 1rem;
    }
    .hero h1 { color:white; margin:0; font-size:2rem; }
    .hero p { color:#e5efff; margin:.4rem 0 0 0; }
    .metric-card {
        background:white; border:1px solid #e5eaf2; border-radius:14px;
        padding:1rem 1.1rem; min-height:100px;
    }
    .metric-label { color:#64748b; font-size:.86rem; }
    .metric-value { color:#12345a; font-size:1.55rem; font-weight:750; }
    div.stButton > button[kind="primary"] {
        background:#2563eb; border-color:#2563eb; border-radius:10px;
        font-weight:650; min-height:2.8rem;
    }
    .hint { color:#64748b; font-size:.9rem; }
    </style>
    """,
    unsafe_allow_html=True,
)

# ---------- Helpers ----------
def prompt_to_columns(prompt: str):
    """Simple starter parser. This is NOT AI; it only recognizes explicit column lists."""
    patterns = [
        r"(?:columns?|fields?|include|with|there should be)\s*[:\-]\s*(.+)",
        r"(?:where|যেখানে)\s*(?:there (?:are|will be))?\s*(.+)",
    ]
    candidate = ""
    for pattern in patterns:
        match = re.search(pattern, prompt, flags=re.IGNORECASE | re.DOTALL)
        if match:
            candidate = match.group(1).split("\n")[0]
            break
    if not candidate:
        # Recognize common "there will be A, B, C" / "with A, B, C" prompts.
        match = re.search(r"\b(?:with|include|including)\s+(.+)", prompt, flags=re.I)
        if match:
            candidate = match.group(1).split("\n")[0]
    if not candidate:
        return []
    candidate = re.sub(r"[.。].*$", "", candidate)
    parts = re.split(r",|;|\band\b|\&", candidate, flags=re.I)
    cols = [p.strip(" \t:-•") for p in parts]
    cols = [c for c in cols if 1 < len(c) < 45]
    return list(dict.fromkeys(cols))[:20]


def make_starter_dataframe(prompt: str, template: str):
    if template == "Employee Details":
        return pd.DataFrame(columns=[
            "Employee ID", "Employee Name", "Department",
            "Designation", "Phone", "Joining Date"
        ])
    if template == "Salary Calculator":
        return pd.DataFrame(columns=[
            "Employee ID", "Employee Name", "Basic Salary",
            "HRA", "Other Allowance", "Gross Salary",
            "Deduction", "Net Salary"
        ])
    if template == "Attendance Sheet":
        return pd.DataFrame(columns=[
            "Employee ID", "Employee Name", "Working Days",
            "Present Days", "Absent", "Attendance %"
        ])
    if template == "Student Marks":
        return pd.DataFrame(columns=[
            "Student ID", "Student Name", "Subject",
            "Total Marks", "Obtained Marks", "Percentage"
        ])
    cols = prompt_to_columns(prompt)
    if not cols:
        cols = ["ID", "Name", "Category", "Date", "Notes"]
    return pd.DataFrame(columns=cols)


def parse_uploaded_file(uploaded_file):
    name = uploaded_file.name.lower()
    raw = uploaded_file.getvalue()
    if name.endswith(".csv"):
        return pd.read_csv(io.BytesIO(raw))
    if name.endswith((".xlsx", ".xls")):
        return pd.read_excel(io.BytesIO(raw))
    raise ValueError("এই প্রথম সংস্করণে CSV এবং XLSX ফাইল সমর্থিত। PDF ও Image extraction পরের ধাপে যোগ হবে।")


def style_workbook(df: pd.DataFrame, sheet_name="Generated Data"):
    output = io.BytesIO()
    safe_df = df.copy().fillna("")
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        safe_df.to_excel(writer, index=False, sheet_name=sheet_name[:31])
        ws = writer.sheets[sheet_name[:31]]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        ws.sheet_view.showGridLines = False
        ws.row_dimensions[1].height = 28

        header_fill = PatternFill("solid", fgColor="17365D")
        header_font = Font(name="Aptos", bold=True, color="FFFFFF", size=11)
        thin_blue = Side(style="thin", color="D7E2F0")
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = Border(left=thin_blue, right=thin_blue, top=thin_blue, bottom=thin_blue)

        for row_idx in range(2, ws.max_row + 1):
            ws.row_dimensions[row_idx].height = 22
            for cell in ws[row_idx]:
                cell.font = Font(name="Aptos", size=10, color="243247")
                cell.alignment = Alignment(vertical="center", wrap_text=True)
                cell.border = Border(left=thin_blue, right=thin_blue, top=thin_blue, bottom=thin_blue)
                if row_idx % 2 == 0:
                    cell.fill = PatternFill("solid", fgColor="F1F6FC")

        for col_idx, column_cells in enumerate(ws.columns, start=1):
            values = [str(c.value) if c.value is not None else "" for c in column_cells]
            longest = max([len(v) for v in values] + [10])
            ws.column_dimensions[get_column_letter(col_idx)].width = min(max(longest + 3, 14), 34)

        if ws.max_row >= 2 and ws.max_column >= 1:
            ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"
            table = Table(displayName="GeneratedTable", ref=ref)
            table.tableStyleInfo = TableStyleInfo(
                name="TableStyleMedium2",
                showFirstColumn=False,
                showLastColumn=False,
                showRowStripes=True,
                showColumnStripes=False,
            )
            ws.add_table(table)

        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.page_setup.orientation = "landscape"
    output.seek(0)
    return output.getvalue()


# ---------- Session state ----------
if "dataframe" not in st.session_state:
    st.session_state.dataframe = pd.DataFrame(columns=["ID", "Name", "Date", "Amount", "Notes"])
if "file_name" not in st.session_state:
    st.session_state.file_name = "AI_Generated_Excel.xlsx"

# ---------- Sidebar ----------
with st.sidebar:
    st.markdown("## 📊 AnalyzeData.io")
    st.caption("AI Excel Generator")
    page = st.radio(
        "NAVIGATION",
        ["Dashboard", "AI Excel Generator", "Templates", "About"],
        label_visibility="visible",
    )
    st.divider()
    st.caption("FIRST RELEASE")
    st.markdown("**Excel generation:** Available")
    st.markdown("**AI connection:** Coming in next step")
    st.markdown("**PDF / Image extraction:** Coming in next step")

# ---------- Dashboard ----------
if page == "Dashboard":
    st.markdown(
        """
        <div class="hero">
          <h1>Turn your data into professional Excel files.</h1>
          <p>Describe the spreadsheet you need, start from a template, or import an existing CSV/XLSX file.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    c1, c2, c3 = st.columns(3)
    for col, label, value in [
        (c1, "Workbook builder", "Ready"),
        (c2, "Editable preview", "Included"),
        (c3, "AI connection", "Next step"),
    ]:
        with col:
            st.markdown(
                f'<div class="metric-card"><div class="metric-label">{label}</div><div class="metric-value">{value}</div></div>',
                unsafe_allow_html=True,
            )
    st.write("")
    st.subheader("Start creating")
    left, right = st.columns(2)
    with left:
        st.markdown("### ✨ Build from a prompt")
        st.write("Describe the columns and layout you want.")
        if st.button("Open AI Excel Generator", type="primary", use_container_width=True):
            st.session_state["nav_target"] = "AI Excel Generator"
            st.info("বাম পাশের Navigation থেকে AI Excel Generator নির্বাচন করুন।")
    with right:
        st.markdown("### 📁 Import existing data")
        st.write("Upload a CSV or Excel file and edit it before downloading.")
        st.info("CSV এবং XLSX import এই সংস্করণে চালু আছে।")

# ---------- Generator ----------
elif page == "AI Excel Generator":
    st.title("AI Excel Generator")
    st.caption("Create a workbook structure, edit the preview, and download a formatted XLSX file.")
    st.text_area(
        "Describe the Excel file you want",
        key="prompt_text",
        height=120,
        placeholder="Example: Create an employee sheet with columns: Employee ID, Name, Department, Phone, Joining Date",
    )
    st.markdown('<p class="hint">Tip: স্পষ্টভাবে column-এর নাম লিখুন—যেমন “columns: Name, Phone, Department, Joining Date”.</p>', unsafe_allow_html=True)

    col_a, col_b = st.columns([1, 1])
    with col_a:
        template = st.selectbox(
            "Start with a template",
            ["Custom / Universal", "Employee Details", "Salary Calculator", "Attendance Sheet", "Student Marks"],
        )
    with col_b:
        upload = st.file_uploader(
            "Or upload source data (CSV / XLSX)",
            type=["csv", "xlsx", "xls"],
            help="PDF ও image extraction পরবর্তী সংস্করণে যোগ করা হবে।",
        )

    action1, action2 = st.columns([1, 1])
    with action1:
        if st.button("Generate spreadsheet structure", type="primary", use_container_width=True):
            st.session_state.dataframe = make_starter_dataframe(
                st.session_state.get("prompt_text", ""), template
            )
            st.session_state.file_name = "AI_Generated_Excel.xlsx"
            st.success("Spreadsheet structure তৈরি হয়েছে। নিচে data edit করতে পারবেন।")
    with action2:
        if upload is not None and st.button("Load uploaded file", use_container_width=True):
            try:
                st.session_state.dataframe = parse_uploaded_file(upload)
                st.session_state.file_name = re.sub(r"\.[^.]+$", "", upload.name) + "_formatted.xlsx"
                st.success(f"{upload.name} সফলভাবে load হয়েছে।")
            except Exception as exc:
                st.error(str(exc))

    st.divider()
    st.subheader("Excel Preview & Editor")
    st.caption("নিচের টেবিলে সরাসরি data লিখুন, row যোগ করুন, বা column-এর তথ্য পরিবর্তন করুন।")
    edited_df = st.data_editor(
        st.session_state.dataframe,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        key="excel_editor",
    )
    st.session_state.dataframe = edited_df

    if len(edited_df.columns) > 0:
        download_bytes = style_workbook(edited_df)
        st.download_button(
            "⬇️ Download Professional Excel",
            data=download_bytes,
            file_name=st.session_state.file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
    st.caption("নোট: এই প্রথম সংস্করণে prompt থেকে column structure তৈরির জন্য সাধারণ rule-based parsing ব্যবহার করা হয়েছে; এটি AI নয়। Gemini API যুক্ত হলে প্রকৃত AI prompt understanding, formula suggestions এবং PDF/Image extraction যোগ হবে।")

# ---------- Templates ----------
elif page == "Templates":
    st.title("Ready-made Excel Templates")
    st.write("যেকোনো template বেছে নিয়ে Generator পেজে editable structure তৈরি করতে পারবেন।")
    templates = {
        "Employee Details": "Employee ID, Employee Name, Department, Designation, Phone, Joining Date",
        "Salary Calculator": "Employee ID, Employee Name, Basic Salary, HRA, Other Allowance, Gross Salary, Deduction, Net Salary",
        "Attendance Sheet": "Employee ID, Employee Name, Working Days, Present Days, Absent, Attendance %",
        "Student Marks": "Student ID, Student Name, Subject, Total Marks, Obtained Marks, Percentage",
    }
    for name, cols in templates.items():
        with st.expander(name):
            st.write(cols)
            st.caption("Generator পেজে এই template নির্বাচন করে structure তৈরি করুন।")

# ---------- About ----------
else:
    st.title("About AnalyzeData.io")
    st.write(
        "AnalyzeData.io-এর লক্ষ্য হলো text prompt বা uploaded data থেকে ব্যবহারযোগ্য, "
        "সুন্দরভাবে formatted Excel workbook তৈরি করা। এই প্রথম release-এ dashboard, "
        "template structure, CSV/XLSX import, editable preview এবং XLSX export রয়েছে।"
    )
    st.warning(
        "Gemini AI, PDF parsing, image OCR এবং automatic formula generation এখনো যুক্ত করা হয়নি। "
        "এসব feature বাস্তবে integrate ও test করার পরেই চালু বলে দেখানো হবে।"
    )
