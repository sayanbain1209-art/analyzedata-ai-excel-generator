import io
import os
import re
import csv
from datetime import date
from collections import OrderedDict

import pandas as pd
import streamlit as st
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter


# =========================================================
# AnalyzeData.io — Universal Official Work & Excel System
# =========================================================
st.set_page_config(
    page_title="AnalyzeData.io | Universal Office & Excel",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    :root { color-scheme: light; }
    .stApp { background: #f4f7fc; color: #172b45; }
    [data-testid="stSidebar"] { background: #102b4c; }
    [data-testid="stSidebar"] * { color: #f4f8ff !important; }
    [data-testid="stSidebar"] [data-testid="stRadio"] label {
        background: rgba(255,255,255,.08);
        border: 1px solid rgba(255,255,255,.18);
        border-radius: 10px; padding: .55rem .7rem; margin: .18rem 0;
    }
    [data-testid="stSidebar"] [data-testid="stRadio"] label:has(input:checked) {
        background: #2563eb; border-color: #9bc0ff;
    }
    .hero {
        padding: 2rem; border-radius: 20px; color: white;
        background: linear-gradient(120deg, #102b4c, #2563eb);
        margin-bottom: 1rem;
    }
    .hero h1, .hero p { color: white !important; }
    .metric-card {
        background: white; border: 1px solid #dce5f1;
        border-radius: 14px; padding: 1rem; min-height: 95px;
    }
    .metric-label { color: #526783; font-size: .85rem; }
    .metric-value { color: #102b4c; font-size: 1.5rem; font-weight: 750; }
    h1, h2, h3, p, label { color: #172b45; }
    .stTextArea textarea, .stTextInput input {
        color: #172b45 !important; background: white !important;
    }
    .stButton button, .stDownloadButton button {
        background: #2563eb; color: white; border: 1px solid #2563eb;
        border-radius: 10px; font-weight: 650; min-height: 2.65rem;
    }
    .stButton button:hover, .stDownloadButton button:hover {
        background: #1748b8; color: white;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# ------------------------------
# DATA HELPERS
# ------------------------------
def unique_headers(values):
    """Keep every source column; give blank/duplicate headings safe names."""
    result, used = [], {}
    for index, value in enumerate(values, start=1):
        name = "" if value is None else str(value).strip()
        if not name or name.lower().startswith("unnamed:"):
            name = f"Column_{index}"
        count = used.get(name.casefold(), 0) + 1
        used[name.casefold()] = count
        if count > 1:
            name = f"{name}_{count}"
        result.append(name)
    return result


def clean_df(df):
    if df is None:
        return pd.DataFrame()
    df = df.copy()
    df.columns = unique_headers(list(df.columns))
    df = df.dropna(how="all")
    return df.reset_index(drop=True)


def detect_header_row(raw_df, max_rows=15):
    if raw_df is None or raw_df.empty:
        return 0
    best_row, best_score = 0, -1
    for row_index in range(min(len(raw_df), max_rows)):
        row = raw_df.iloc[row_index].tolist()
        values = [v for v in row if not pd.isna(v) and str(v).strip() != ""]
        if not values:
            continue
        text_count = sum(isinstance(v, str) for v in values)
        unique_count = len({str(v).strip().casefold() for v in values})
        nonempty = len(values)
        score = nonempty * 2 + (text_count / nonempty) * 3 + (unique_count / nonempty)
        score -= row_index * 0.025
        if score > best_score:
            best_row, best_score = row_index, score
    return best_row


def dataframe_from_raw(raw_df):
    if raw_df is None or raw_df.empty:
        return pd.DataFrame()
    raw_df = raw_df.dropna(how="all")
    if raw_df.empty:
        return pd.DataFrame()
    header_row = detect_header_row(raw_df)
    headers = unique_headers(raw_df.iloc[header_row].tolist())
    body = raw_df.iloc[header_row + 1:].copy()
    body.columns = headers
    return clean_df(body)


def parse_pasted_data(raw):
    raw = (raw or "").strip()
    if not raw:
        return pd.DataFrame()
    sample = "\n".join(raw.splitlines()[:8])
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",\t;|")
        df = pd.read_csv(io.StringIO(raw), sep=dialect.delimiter, engine="python")
        if len(df.columns) > 1:
            return clean_df(df)
    except Exception:
        pass
    try:
        df = pd.read_csv(io.StringIO(raw), sep=None, engine="python")
        if len(df.columns) > 1:
            return clean_df(df)
    except Exception:
        pass
    raise ValueError("প্রথম লাইনে column heading দিন এবং column-গুলো কমা বা Tab দিয়ে আলাদা করুন।")


def read_uploaded_workbook(uploaded_file):
    content = uploaded_file.getvalue()
    filename = uploaded_file.name.lower()
    result = OrderedDict()

    if filename.endswith((".xlsx", ".xlsm", ".xls")):
        excel = pd.ExcelFile(io.BytesIO(content))
        for sheet_name in excel.sheet_names:
            raw = pd.read_excel(io.BytesIO(content), sheet_name=sheet_name, header=None)
            result[str(sheet_name)] = dataframe_from_raw(raw)
    elif filename.endswith(".tsv"):
        raw = pd.read_csv(io.BytesIO(content), sep="\t", header=None)
        result["Data"] = dataframe_from_raw(raw)
    else:
        raw = pd.read_csv(io.BytesIO(content), sep=None, engine="python", header=None)
        result["Data"] = dataframe_from_raw(raw)

    return result


# ------------------------------
# SALARY CALCULATIONS
# ------------------------------
def salary_calculations(df):
    out = clean_df(df)
    lookup = {str(c).strip().casefold(): c for c in out.columns}

    def find(*names):
        for name in names:
            if name.casefold() in lookup:
                return lookup[name.casefold()]
        return None

    basic = find("basic salary", "basic", "basic pay", "basic wages")
    if basic is None:
        return out

    hra = find("hra", "house rent allowance", "rent allowance")
    allowance = find("other allowance", "allowance", "other allowances", "special allowance", "travel allowance")
    deduction = find("deduction", "deductions", "total deduction")

    def numbers(column):
        if column is None:
            return pd.Series(0.0, index=out.index)
        return pd.to_numeric(
            out[column].astype(str)
            .str.replace(",", "", regex=False)
            .str.replace("₹", "", regex=False)
            .str.strip(),
            errors="coerce",
        ).fillna(0)

    out["Gross Salary"] = numbers(basic) + numbers(hra) + numbers(allowance)
    out["Net Salary"] = out["Gross Salary"] - numbers(deduction)
    return out


# ------------------------------
# DATA PROFILE / QUALITY
# ------------------------------
def data_profile(df):
    if df.empty:
        return pd.DataFrame()
    rows = []
    for column in df.columns:
        series = df[column]
        rows.append({
            "Column": str(column),
            "Data Type": str(series.dtype),
            "Filled Values": int(series.notna().sum()),
            "Missing Values": int(series.isna().sum()),
            "Unique Values": int(series.nunique(dropna=True)),
        })
    return pd.DataFrame(rows)


def make_excel(sheets):
    """Create a formatted multi-sheet Excel workbook with robust formulas & styling."""
    if isinstance(sheets, pd.DataFrame):
        sheets = {"Data": sheets}

    output = io.BytesIO()
    workbook = Workbook()
    workbook.remove(workbook.active)

    used_sheet_names = set()
    for original_name, source_df in sheets.items():
        df = clean_df(source_df)
        base_name = re.sub(r"[\[\]:*?/\\]", "_", str(original_name)).strip() or "Data"
        base_name = base_name[:31]
        sheet_name = base_name
        suffix = 2
        while sheet_name.casefold() in used_sheet_names:
            tail = f"_{suffix}"
            sheet_name = base_name[:31 - len(tail)] + tail
            suffix += 1
        used_sheet_names.add(sheet_name.casefold())

        ws = workbook.create_sheet(title=sheet_name)
        ws.sheet_state = "visible"
        ws.sheet_view.showGridLines = True
        ws.sheet_view.showRowColHeaders = True
        ws.sheet_view.zoomScale = 100

        if len(df.columns) == 0:
            ws.cell(row=1, column=1, value="No columns")
            continue

        headers = unique_headers(list(df.columns))

        # Detect Salary columns to inject live Excel formulas
        col_indices = {h.casefold(): idx for idx, h in enumerate(headers, start=1)}
        basic_col = col_indices.get("basic salary", col_indices.get("basic"))
        hra_col = col_indices.get("hra")
        allow_col = col_indices.get("other allowance", col_indices.get("allowance"))
        ded_col = col_indices.get("deduction", col_indices.get("deductions"))
        gross_col = col_indices.get("gross salary")
        net_col = col_indices.get("net salary")

        # Header Styling
        header_fill = PatternFill(fill_type="solid", fgColor="1B365D")
        header_font = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
        header_border = Border(
            top=Side(style="thin", color="1B365D"),
            bottom=Side(style="medium", color="0D1B2A"),
            left=Side(style="thin", color="415A77"),
            right=Side(style="thin", color="415A77"),
        )

        for col_num, heading in enumerate(headers, start=1):
            cell = ws.cell(row=1, column=col_num, value=str(heading))
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = header_border

        # Body rows & cell formatting
        thin_border = Border(
            left=Side(style="thin", color="E0E0E0"),
            right=Side(style="thin", color="E0E0E0"),
            top=Side(style="thin", color="E0E0E0"),
            bottom=Side(style="thin", color="E0E0E0"),
        )
        stripe_fill = PatternFill(fill_type="solid", fgColor="F8FAFC")

        for row_num, record in enumerate(df.itertuples(index=False, name=None), start=2):
            for col_num, value in enumerate(record, start=1):
                cell = ws.cell(row=row_num, column=col_num)

                # Inject dynamic Excel Formulas for Payroll
                if gross_col and col_num == gross_col and basic_col:
                    add_parts = [f"{get_column_letter(c)}{row_num}" for c in [basic_col, hra_col, allow_col] if c]
                    cell.value = f"={' + '.join(add_parts)}"
                elif net_col and col_num == net_col and gross_col:
                    if ded_col:
                        cell.value = f"={get_column_letter(gross_col)}{row_num} - {get_column_letter(ded_col)}{row_num}"
                    else:
                        cell.value = f"={get_column_letter(gross_col)}{row_num}"
                else:
                    if pd.isna(value):
                        cell.value = None
                    elif hasattr(value, "to_pydatetime"):
                        cell.value = value.to_pydatetime()
                    else:
                        try:
                            # Parse numeric values cleanly
                            num_str = str(value).replace(",", "").replace("₹", "").strip()
                            if num_str.replace(".", "", 1).isdigit():
                                cell.value = float(num_str) if "." in num_str else int(num_str)
                            else:
                                cell.value = value
                        except Exception:
                            cell.value = value

                # Number formatting
                if isinstance(cell.value, (int, float)):
                    cell.number_format = "#,##0.00"
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center")

                if row_num % 2 == 0:
                    cell.fill = stripe_fill
                cell.border = thin_border

        ws.row_dimensions[1].height = 28
        ws.freeze_panes = "A2"

        # Auto column width
        for col_num, heading in enumerate(headers, start=1):
            letter = get_column_letter(col_num)
            max_len = len(str(heading))
            for val in df.iloc[:, col_num - 1].dropna().head(200):
                max_len = max(max_len, len(str(val)))
            ws.column_dimensions[letter].width = min(max(max_len + 4, 12), 35)

    if not workbook.sheetnames:
        workbook.create_sheet("Data")
    workbook.save(output)
    return output.getvalue()


# ------------------------------
# GEMINI AI
# ------------------------------
def gemini_answer(prompt, df):
    try:
        key = st.secrets.get("GEMINI_API_KEY", None)
    except Exception:
        key = None
    key = key or os.getenv("GEMINI_API_KEY")
    if not key:
        return (
            "Gemini এখনো সংযুক্ত নয়। Streamlit Cloud → App settings → "
            "Secrets-এ GEMINI_API_KEY যোগ করুন।"
        )
    from google import genai
    client = genai.Client(api_key=key)
    
    context = (
        "You are an expert Excel & Data Analyst.\n"
        "Data sample (CSV format up to 40 rows):\n"
        + df.head(40).to_csv(index=False)
        + "\n\nGive direct, actionable analysis. If user asks for new columns or formulas, describe them clearly."
    )
    response = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=prompt + "\n\n" + context,
    )
    return response.text or "AI কোনো উত্তর ফেরত দেয়নি।"


# ------------------------------
# OFFICIAL WORK TEMPLATE LIBRARY
# ------------------------------
TEMPLATES = {
    "HR & Employee": {
        "Employee Details": ["Employee ID", "Employee Name", "Department", "Designation", "Joining Date", "Email", "Phone", "Status"],
        "Attendance Register": ["Date", "Employee ID", "Employee Name", "Department", "Status", "In Time", "Out Time", "Working Hours"],
        "Leave Register": ["Leave ID", "Employee ID", "Employee Name", "Leave Type", "From Date", "To Date", "Days", "Approval Status"],
        "Recruitment Tracker": ["Candidate ID", "Candidate Name", "Position", "Department", "Application Date", "Interview Date", "Stage", "Result"],
    },
    "Payroll & Finance": {
        "Employee Salary": ["Employee ID", "Employee Name", "Department", "Basic Salary", "HRA", "Other Allowance", "Deduction", "Gross Salary", "Net Salary"],
        "Cash Book": ["Date", "Voucher No", "Description", "Cash In", "Cash Out", "Balance", "Reference"],
        "Expense Register": ["Date", "Expense ID", "Description", "Category", "Payment Method", "Amount", "Approved By"],
        "Accounts Ledger": ["Date", "Voucher No", "Account Name", "Debit", "Credit", "Narration", "Reference"],
        "Invoice Register": ["Invoice No", "Invoice Date", "Customer", "Description", "Quantity", "Unit Price", "Tax", "Total"],
    },
    "Business & Inventory": {
        "Sales Report": ["Date", "Invoice No", "Customer", "Product", "Category", "Quantity", "Unit Price", "Total Sales"],
        "Purchase Register": ["Date", "Purchase No", "Supplier", "Product", "Quantity", "Unit Cost", "Total Cost", "Payment Status"],
        "Inventory Stock": ["Item Code", "Item Name", "Category", "Opening Stock", "Stock In", "Stock Out", "Closing Stock", "Reorder Level"],
        "Supplier Details": ["Supplier ID", "Supplier Name", "Contact Person", "Phone", "Email", "Address", "Payment Terms"],
        "Customer Details": ["Customer ID", "Customer Name", "Phone", "Email", "Address", "Opening Balance", "Status"],
    },
    "Education": {
        "Student Register": ["Student ID", "Student Name", "Class", "Section", "Roll", "Date of Birth", "Guardian Name", "Phone"],
        "Student Marks": ["Roll", "Student Name", "Subject", "Total Marks", "Marks Obtained", "Percentage", "Grade"],
        "Fee Register": ["Receipt No", "Date", "Student ID", "Student Name", "Fee Type", "Amount Due", "Amount Paid", "Balance"],
    },
    "Healthcare": {
        "Patient Register": ["Patient ID", "Patient Name", "Age", "Phone", "Address", "Registration Date", "Assigned Doctor"],
        "Appointment Register": ["Appointment ID", "Date", "Time", "Patient ID", "Patient Name", "Doctor", "Department", "Status"],
        "Pharmacy Stock": ["Medicine ID", "Medicine Name", "Batch No", "Expiry Date", "Quantity", "Unit Price", "Supplier"],
    },
    "Projects & Administration": {
        "Project Tracker": ["Project ID", "Project Name", "Manager", "Department", "Start Date", "Due Date", "Budget", "Status"],
        "Task Register": ["Task ID", "Task Name", "Assigned To", "Priority", "Start Date", "Due Date", "Status", "Remarks"],
        "Asset Register": ["Asset ID", "Asset Name", "Category", "Location", "Purchase Date", "Cost", "Assigned To", "Condition"],
        "Document Register": ["Document ID", "Document Title", "Document Type", "Department", "Received Date", "Owner", "Status", "Remarks"],
        "Maintenance Log": ["Ticket ID", "Asset/Location", "Issue", "Reported Date", "Assigned To", "Completion Date", "Status"],
    },
    "Logistics & Other Services": {
        "Delivery Register": ["Shipment ID", "Date", "Sender", "Receiver", "Origin", "Destination", "Tracking No", "Delivery Status"],
        "Vehicle Log": ["Date", "Vehicle No", "Driver", "Opening KM", "Closing KM", "Distance", "Fuel", "Purpose"],
        "Hotel Booking": ["Booking ID", "Guest Name", "Check-in", "Check-out", "Room No", "Amount", "Payment Status"],
        "Donation Register": ["Receipt No", "Date", "Donor Name", "Purpose", "Amount", "Payment Method", "Reference"],
    },
}


# ------------------------------
# SESSION STATE
# ------------------------------
if "sheets" not in st.session_state:
    st.session_state.sheets = OrderedDict({"Data": pd.DataFrame()})
if "active_sheet" not in st.session_state:
    st.session_state.active_sheet = "Data"
if "ai_result" not in st.session_state:
    st.session_state.ai_result = ""


def active_df():
    name = st.session_state.active_sheet
    if name not in st.session_state.sheets:
        st.session_state.active_sheet = next(iter(st.session_state.sheets))
        name = st.session_state.active_sheet
    return st.session_state.sheets[name]


# ------------------------------
# SIDEBAR
# ------------------------------
with st.sidebar:
    st.markdown("## 📊 AnalyzeData.io")
    st.caption("Universal Official Work & Excel System")
    st.markdown("---")
    page = st.radio(
        "NAVIGATION",
        [
            "🏠 Dashboard",
            "🌐 Universal Data Workspace",
            "📁 Official Templates",
            "ℹ️ About",
        ],
    )
    st.markdown("---")
    st.markdown("### CORE FEATURES")
    st.markdown(
        """
        🟢 **Clean & Visible Headings**
        
        🟢 **Excel Dynamic Formulas**
        
        🟢 **Formatted Numbers & Grid**
        
        🟢 **Multi-sheet CSV / XLSX**
        
        🟢 **Gemini AI Integration**
        """
    )


# ------------------------------
# DASHBOARD
# ------------------------------
if page == "🏠 Dashboard":
    st.markdown(
        """
        <div class="hero">
            <h1>AnalyzeData.io</h1>
            <p>Universal Official Work Management & Excel Analysis</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    all_frames = list(st.session_state.sheets.values())
    total_rows = sum(len(frame) for frame in all_frames)
    total_columns = sum(len(frame.columns) for frame in all_frames)
    col1, col2, col3, col4 = st.columns(4)
    cards = [
        (col1, "Worksheets", len(all_frames)),
        (col2, "Total records", total_rows),
        (col3, "Total columns", total_columns),
        (col4, "Templates", sum(len(v) for v in TEMPLATES.values())),
    ]
    for column, label, value in cards:
        with column:
            st.markdown(
                f'<div class="metric-card"><div class="metric-label">{label}</div>'
                f'<div class="metric-value">{value:,}</div></div>',
                unsafe_allow_html=True,
            )

    st.subheader("একটি কাজ বেছে নিন")
    left, right = st.columns(2)
    with left:
        st.info("🌐 **Universal Data Workspace**\n\nফাইল upload, paste, edit, data quality check ও export করুন।")
    with right:
        st.info("📁 **Official Templates**\n\nবিভিন্ন ক্যাটাগরির ready-made structure নিয়ে কাজ শুরু করুন।")


# ------------------------------
# UNIVERSAL DATA WORKSPACE
# ------------------------------
elif page == "🌐 Universal Data Workspace":
    st.title("🌐 Universal Data Workspace")
    paste_tab, upload_tab = st.tabs(["📝 Paste data", "📤 Upload workbook"])

    with paste_tab:
        raw = st.text_area(
            "টেবিল ডেটা পেস্ট করুন (প্রথম লাইনে column heading রাখুন)",
            height=140,
            placeholder="Employee ID,Employee Name,Department,Basic Salary,HRA,Deduction\nEMP101,Rahul Das,Sales,18000,3000,500",
            key="paste_data_area",
        )
        if st.button("Import pasted data", use_container_width=True):
            try:
                parsed = parse_pasted_data(raw)
                if parsed.empty:
                    st.warning("আগে কিছু ডেটা পেস্ট করুন।")
                else:
                    st.session_state.sheets = OrderedDict({"Data": parsed})
                    st.session_state.active_sheet = "Data"
                    st.success(f"{len(parsed)}টি row এবং {len(parsed.columns)}টি column import হয়েছে।")
            except Exception as error:
                st.error(str(error))

    with upload_tab:
        uploaded_file = st.file_uploader("Excel/CSV ফাইল আপলোড করুন", type=["csv", "xlsx", "xlsm", "xls", "tsv"])
        if uploaded_file is not None:
            if st.button("Import workbook", use_container_width=True):
                try:
                    imported = read_uploaded_workbook(uploaded_file)
                    st.session_state.sheets = imported
                    st.session_state.active_sheet = next(iter(imported))
                    st.success(f"{len(imported)}টি worksheet import হয়েছে।")
                except Exception as error:
                    st.error(f"ফাইল পড়া যায়নি: {error}")

    if st.session_state.sheets:
        sheet_names = list(st.session_state.sheets.keys())
        if len(sheet_names) > 1:
            selected = st.selectbox("কোন worksheet নিয়ে কাজ করবেন?", sheet_names)
            st.session_state.active_sheet = selected

        current_name = st.session_state.active_sheet
        current = active_df()

        st.divider()
        st.subheader(f"📊 {current_name} — Preview & Editor")
        if not current.empty or len(current.columns) > 0:
            edited = st.data_editor(current, num_rows="dynamic", use_container_width=True, hide_index=True)
            st.session_state.sheets[current_name] = clean_df(edited)
            current = active_df()

            col_a, col_b = st.columns(2)
            with col_a:
                if st.button("🧮 Auto Calculate Gross & Net Salary", use_container_width=True):
                    result = salary_calculations(current)
                    st.session_state.sheets[current_name] = result
                    st.success("Gross ও Net Salary হিসাব সম্পন্ন!")
                    st.rerun()
            with col_b:
                try:
                    export_sheets = OrderedDict(
                        (name, salary_calculations(frame))
                        for name, frame in st.session_state.sheets.items()
                    )
                    excel_bytes = make_excel(export_sheets)
                    st.download_button(
                        "⬇️ Download Professional Excel Sheet",
                        data=excel_bytes,
                        file_name=f"AnalyzedData_{date.today().isoformat()}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
                except Exception as error:
                    st.error(f"Excel export error: {error}")

            st.divider()
            st.subheader("🤖 Gemini AI Data Assistant")
            question = st.text_area("AI-কে কোনো প্রশ্ন বা হিসাব করতে বলুন:", placeholder="উদা: কোনো অসঙ্গতি আছে কি না বা ডিপার্টমেন্ট অনুযায়ী গড় স্যালারি কত?")
            if st.button("Ask AI"):
                if question.strip():
                    with st.spinner("AI বিশ্লেষণ করছে..."):
                        st.session_state.ai_result = gemini_answer(question, current)
                else:
                    st.warning("একটি প্রশ্ন লিখুন।")
            if st.session_state.ai_result:
                st.markdown(st.session_state.ai_result)


# ------------------------------
# OFFICIAL TEMPLATES
# ------------------------------
elif page == "📁 Official Templates":
    st.title("📁 Official Work Templates")
    category = st.selectbox("কাজের বিভাগ", list(TEMPLATES.keys()))
    chosen = st.selectbox("Template নির্বাচন করুন", list(TEMPLATES[category].keys()))
    columns = TEMPLATES[category][chosen]
    st.subheader(chosen)
    st.dataframe(pd.DataFrame(columns=columns), use_container_width=True, hide_index=True)
    if st.button("Use this template", use_container_width=True):
        st.session_state.sheets[chosen] = pd.DataFrame(columns=columns)
        st.session_state.active_sheet = chosen
        st.success(f"'{chosen}' টেমপ্লেট লোড হয়েছে! Universal Data Workspace-এ গিয়ে ডেটা এন্ট্রি করুন।")


# ------------------------------
# ABOUT
# ------------------------------
else:
    st.title("ℹ️ About AnalyzeData.io")
    st.write("Universal Official Work Management & Excel Analysis System.")