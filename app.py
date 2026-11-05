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
from openpyxl.worksheet.table import Table, TableStyleInfo


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
    """
    Estimate a header row from the first rows of a worksheet.
    This is a best-effort heuristic; users can inspect and edit headings later.
    """
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
        # Headers are commonly text-heavy and contain distinct labels.
        score = nonempty * 2 + (text_count / nonempty) * 3 + (unique_count / nonempty)
        # Prefer earlier rows when scores are close.
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
    """Read every sheet in an Excel workbook, not only the first sheet."""
    content = uploaded_file.getvalue()
    filename = uploaded_file.name.lower()
    result = OrderedDict()

    if filename.endswith((".xlsx", ".xlsm", ".xls")):
        excel = pd.ExcelFile(io.BytesIO(content))
        for sheet_name in excel.sheet_names:
            raw = pd.read_excel(
                io.BytesIO(content), sheet_name=sheet_name, header=None
            )
            result[str(sheet_name)] = dataframe_from_raw(raw)
    elif filename.endswith(".tsv"):
        raw = pd.read_csv(io.BytesIO(content), sep="\t", header=None)
        result["Data"] = dataframe_from_raw(raw)
    else:
        raw = pd.read_csv(
            io.BytesIO(content), sep=None, engine="python", header=None
        )
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
    allowance = find(
        "other allowance", "allowance", "other allowances",
        "special allowance", "travel allowance"
    )
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


def export_headers_for_sheet(sheet_name, df):
    """Return meaningful, non-empty headings for every exported column."""
    headers = unique_headers(list(df.columns))
    placeholders = all(
        str(h).strip().casefold() in {f"column_{i}".casefold(), f"unnamed: {i-1}".casefold(), ""}
        for i, h in enumerate(headers, start=1)
    )
    if not placeholders:
        return headers

    # Reuse the matching official template when the worksheet name identifies it.
    for category_templates in TEMPLATES.values():
        for template_name, template_headers in category_templates.items():
            if str(template_name).strip().casefold() == str(sheet_name).strip().casefold():
                if len(template_headers) == len(headers):
                    return list(template_headers)

    # Never leave blank/Unnamed headings in an exported workbook.
    return [f"Field {i}" for i in range(1, len(headers) + 1)]


def make_excel(sheets):
    """
    Export every worksheet with its subject-specific column headings in Excel row 1.
    Data always begins in row 2. No fixed Employee-only headings are used.
    """
    if isinstance(sheets, pd.DataFrame):
        sheets = {"Data": sheets}

    output = io.BytesIO()
    workbook = Workbook()
    workbook.remove(workbook.active)

    used_sheet_names = set()
    for original_name, source_df in sheets.items():
        df = clean_df(source_df)
        base_name = re.sub(r"[\\[\\]:*?/\\\\]", "_", str(original_name)).strip() or "Data"
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
        ws.sheet_view.zoomScale = 90

        if len(df.columns) == 0:
            ws.cell(row=1, column=1, value="No columns")
            ws.freeze_panes = "A2"
            continue

        # Use the actual column names for this dataset, whatever its subject is.
        headers = export_headers_for_sheet(original_name, df)
        headers = [str(value).strip() if value is not None else "" for value in headers]
        headers = [
            heading if heading else f"Field {index}"
            for index, heading in enumerate(headers, start=1)
        ]

        # Write row 1 explicitly before writing any records.
        for col_num, heading in enumerate(headers, start=1):
            cell = ws.cell(row=1, column=col_num)
            cell.value = heading
            # Dark text + light fill remains readable in mobile spreadsheet viewers.
            cell.font = Font(name="Arial", bold=True, color="17365D", size=11)
            cell.fill = PatternFill(fill_type="solid", fgColor="D9EAF7")
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = Border(
                top=Side(style="thin", color="7F9DB9"),
                bottom=Side(style="medium", color="1F4E78"),
                left=Side(style="thin", color="B8CDE0"),
                right=Side(style="thin", color="B8CDE0"),
            )

        # All data records start at row 2; preserve the dataset's own column order.
        for row_num, record in enumerate(df.itertuples(index=False, name=None), start=2):
            for col_num, value in enumerate(record, start=1):
                if pd.isna(value):
                    value = None
                if hasattr(value, "to_pydatetime"):
                    value = value.to_pydatetime()
                ws.cell(row=row_num, column=col_num, value=value)

        # Give the heading row enough height and each column enough width
        # so subject-specific headings are comfortably readable.
        ws.row_dimensions[1].height = 36
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{max(ws.max_row, 1)}"

        for col_num, heading in enumerate(headers, start=1):
            letter = get_column_letter(col_num)
            longest = max(
                (len(str(v)) for v in df.iloc[:, col_num - 1].dropna().head(500)),
                default=0,
            )
            # Prioritize the full heading; allow wider columns for longer names.
            heading_width = len(str(heading)) + 6
            ws.column_dimensions[letter].width = min(
                max(heading_width, longest + 2, 16), 42
            )

        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            for cell in row:
                cell.alignment = Alignment(vertical="center", wrap_text=False)
                cell.border = Border(bottom=Side(style="thin", color="D8E3F0"))

        # Verify every column has a visible, non-empty heading in row 1.
        for col_num in range(1, len(headers) + 1):
            if not str(ws.cell(row=1, column=col_num).value or "").strip():
                raise ValueError(
                    f"Excel export validation failed: blank heading in {sheet_name}, column {col_num}."
                )

    if not workbook.sheetnames:
        workbook.create_sheet("Data")
    workbook.save(output)
    excel_bytes = output.getvalue()

    # Re-open the actual XLSX bytes and verify row 1 + row 2 placement.
    check = load_workbook(io.BytesIO(excel_bytes), read_only=False, data_only=False)
    for ws in check.worksheets:
        if ws.max_column > 0:
            for col_num in range(1, ws.max_column + 1):
                value = ws.cell(row=1, column=col_num).value
                if value is None or not str(value).strip():
                    check.close()
                    raise ValueError(
                        f"Excel export validation failed: blank heading in {ws.title}, column {col_num}."
                    )
            if ws.freeze_panes != "A2":
                check.close()
                raise ValueError(
                    f"Excel export validation failed: data/header boundary is incorrect in {ws.title}."
                )
    check.close()
    return excel_bytes


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
            "Secrets-এ GEMINI_API_KEY যোগ করুন। API key app.py-তে লিখবেন না।"
        )
    from google import genai
    client = genai.Client(api_key=key)
    context = "\n\nData sample (up to 40 rows):\n" + df.head(40).to_csv(index=False)
    response = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=prompt + context,
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
        "Employee Salary": ["Employee ID", "Employee Name", "Department", "Basic Salary", "HRA", "Other Allowance", "Deduction"],
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
        "Legal Case Register": ["Case No", "Court/Authority", "Parties", "Case Type", "Filing Date", "Next Hearing", "Responsible Person", "Status"],
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
if "import_message" not in st.session_state:
    st.session_state.import_message = ""


def active_df():
    name = st.session_state.active_sheet
    if name not in st.session_state.sheets:
        st.session_state.active_sheet = next(iter(st.session_state.sheets))
        name = st.session_state.active_sheet
    return st.session_state.sheets[name]


# ------------------------------
# SIDEBAR NAVIGATION
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
        🟢 **CSV / XLSX Import**

        🟢 **Multi-sheet workbook support**

        🟢 **Automatic heading detection**

        🟢 **Data quality & analysis**

        🟢 **Professional Excel export**

        🟢 **Office work templates**

        🟡 **Gemini AI** — API key required
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
            <p>Import · Understand · Edit · Analyze · Create reports · Export</p>
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
    st.write(
        "যেকোনো supported Excel/CSV workbook import করুন, অথবা Official Templates "
        "থেকে HR, Finance, Business, Education, Healthcare, Project ও অন্যান্য কাজের "
        "জন্য ready-made structure তৈরি করুন।"
    )
    left, right = st.columns(2)
    with left:
        st.info("🌐 **Universal Data Workspace**\n\nফাইল upload, paste, edit, data quality check ও analysis করুন।")
    with right:
        st.info("📁 **Official Templates**\n\nবিভিন্ন বিভাগের কাজের জন্য ready-made column headings নিন।")

    if total_rows:
        st.subheader("বর্তমান Workbook")
        for sheet_name, frame in st.session_state.sheets.items():
            st.write(f"**{sheet_name}** — {len(frame)} rows × {len(frame.columns)} columns")
            if not frame.empty:
                st.dataframe(frame.head(5), use_container_width=True, hide_index=True)


# ------------------------------
# UNIVERSAL DATA WORKSPACE
# ------------------------------
elif page == "🌐 Universal Data Workspace":
    st.title("🌐 Universal Data Workspace")
    st.caption(
        "যেকোনো অফিসিয়াল বা business dataset import, inspect, edit, analyze এবং export করুন।"
    )

    paste_tab, upload_tab = st.tabs(["📝 Paste data", "📤 Upload workbook"])

    with paste_tab:
        raw = st.text_area(
            "টেবিল ডেটা পেস্ট করুন (প্রথম লাইনে column heading রাখুন)",
            height=160,
            placeholder=(
                "Employee ID,Employee Name,Department,Basic Salary,HRA,Deduction\n"
                "EMP101,Rahul Das,Sales,18000,3000,500"
            ),
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
                    st.session_state.ai_result = ""
                    st.success(f"{len(parsed)}টি row এবং {len(parsed.columns)}টি column import হয়েছে।")
            except Exception as error:
                st.error(str(error))

    with upload_tab:
        uploaded_file = st.file_uploader(
            "CSV, XLSX, XLSM, XLS বা TSV ফাইল বেছে নিন",
            type=["csv", "xlsx", "xlsm", "xls", "tsv"],
            key="universal_file_upload",
        )
        if uploaded_file is not None:
            st.caption("Excel workbook হলে সব worksheet পড়ার চেষ্টা করা হবে।")
            if st.button("Import workbook", use_container_width=True):
                try:
                    imported = read_uploaded_workbook(uploaded_file)
                    imported = OrderedDict(
                        (name, frame) for name, frame in imported.items()
                        if frame is not None
                    )
                    if not imported:
                        st.warning("ফাইলে কোনো worksheet data পাওয়া যায়নি।")
                    else:
                        st.session_state.sheets = imported
                        st.session_state.active_sheet = next(iter(imported))
                        st.session_state.ai_result = ""
                        details = ", ".join(
                            f"{name}: {len(frame)} rows"
                            for name, frame in imported.items()
                        )
                        st.success(f"{len(imported)}টি worksheet import হয়েছে। {details}")
                        st.info(
                            "Heading auto-detection একটি best-effort ব্যবস্থা। "
                            "নিচের Preview-তে heading পরীক্ষা করে প্রয়োজনে edit করুন।"
                        )
                except Exception as error:
                    st.error(
                        f"ফাইল পড়া যায়নি: {error}. পুরোনো .xls ফাইলের জন্য xlrd package লাগতে পারে।"
                    )

    st.divider()
    st.subheader("নতুন টেবিল তৈরি করুন")
    description = st.text_input(
        "কোন কোন column চান?",
        placeholder="Name, Department, Phone, Joining Date",
        key="new_columns",
    )
    new_sheet_name = st.text_input("Worksheet-এর নাম", value="New Data", key="new_sheet_name")
    if st.button("Create empty structure"):
        columns = [item.strip() for item in re.split(r"[,;|]", description) if item.strip()]
        if columns:
            safe_name = new_sheet_name.strip() or "New Data"
            st.session_state.sheets[safe_name] = pd.DataFrame(columns=unique_headers(columns))
            st.session_state.active_sheet = safe_name
            st.session_state.ai_result = ""
            st.success(f"{safe_name} worksheet তৈরি হয়েছে।")
        else:
            st.warning("কমা দিয়ে column-এর নাম লিখুন।")

    if st.session_state.sheets:
        sheet_names = list(st.session_state.sheets.keys())
        if len(sheet_names) > 1:
            selected = st.selectbox(
                "কোন worksheet নিয়ে কাজ করবেন?",
                sheet_names,
                index=sheet_names.index(st.session_state.active_sheet)
                if st.session_state.active_sheet in sheet_names else 0,
            )
            st.session_state.active_sheet = selected

        current_name = st.session_state.active_sheet
        current = active_df()

        st.divider()
        st.subheader(f"📊 {current_name} — Preview & Editor")
        if current.empty and len(current.columns) == 0:
            st.info("এই worksheet-এ এখনো data নেই। উপরে data paste/upload করুন অথবা template নিন।")
        else:
            st.caption(
                "Excel Export-এ প্রথম সারিতে Column Heading এবং দ্বিতীয় সারি থেকে Data থাকবে।"
            )

            # Explicit heading editor for reliable column-name editing.
            with st.expander("✏️ Column Heading দেখুন / পরিবর্তন করুন", expanded=True):
                heading_values = []
                heading_cols = st.columns(min(len(current.columns), 3))
                for heading_index, old_heading in enumerate(current.columns):
                    with heading_cols[heading_index % len(heading_cols)]:
                        heading_values.append(
                            st.text_input(
                                f"Column {heading_index + 1}",
                                value=str(old_heading),
                                key=f"heading_{current_name}_{heading_index}",
                            )
                        )
                if st.button("Heading Apply করুন", key=f"apply_headings_{current_name}"):
                    cleaned_headings = [str(value).strip() for value in heading_values]
                    if any(not value for value in cleaned_headings):
                        st.error("কোনো Heading ফাঁকা রাখা যাবে না।")
                    else:
                        renamed = current.copy()
                        renamed.columns = unique_headers(cleaned_headings)
                        st.session_state.sheets[current_name] = renamed
                        st.success("Column Heading আপডেট হয়েছে।")
                        st.rerun()

            edited = st.data_editor(
                current,
                num_rows="dynamic",
                use_container_width=True,
                hide_index=True,
                key=f"editor_{current_name}",
            )
            st.session_state.sheets[current_name] = clean_df(edited)
            current = active_df()

            overview_tab, quality_tab, analysis_tab = st.tabs(
                ["📋 Overview", "🧹 Data Quality", "📈 Analysis"]
            )
            with overview_tab:
                c1, c2, c3 = st.columns(3)
                c1.metric("Rows", f"{len(current):,}")
                c2.metric("Columns", f"{len(current.columns):,}")
                c3.metric("Duplicate rows", f"{int(current.duplicated().sum()):,}")
                st.dataframe(current, use_container_width=True, hide_index=True)

            with quality_tab:
                st.write("প্রতি Column-এর data type, missing value ও unique value:")
                st.dataframe(data_profile(current), use_container_width=True, hide_index=True)
                missing_total = int(current.isna().sum().sum()) if not current.empty else 0
                duplicate_total = int(current.duplicated().sum()) if not current.empty else 0
                st.write(f"**মোট missing cell:** {missing_total:,}")
                st.write(f"**Duplicate row:** {duplicate_total:,}")
                if duplicate_total:
                    if st.button("Remove duplicate rows from this worksheet"):
                        st.session_state.sheets[current_name] = current.drop_duplicates().reset_index(drop=True)
                        st.success("Duplicate rows সরানো হয়েছে।")
                        st.rerun()

            with analysis_tab:
                numeric = current.select_dtypes(include="number")
                if numeric.empty:
                    st.info("এই worksheet-এ numeric column পাওয়া যায়নি।")
                else:
                    st.write("Numeric columns-এর summary:")
                    st.dataframe(numeric.describe().T, use_container_width=True)
                    selected_numeric = st.selectbox("একটি numeric column বেছে নিন", list(numeric.columns))
                    st.metric(
                        f"Total — {selected_numeric}",
                        f"{numeric[selected_numeric].sum():,.2f}",
                    )

            st.divider()
            st.subheader("⚙️ Office calculations")
            col_a, col_b, col_c = st.columns(3)
            with col_a:
                if st.button("🧮 Calculate salary", use_container_width=True):
                    result = salary_calculations(current)
                    if result.equals(current):
                        st.warning("Basic Salary column পাওয়া যায়নি। Heading পরীক্ষা করুন।")
                    else:
                        st.session_state.sheets[current_name] = result
                        st.success("Gross Salary এবং Net Salary হিসাব করা হয়েছে।")
                        st.rerun()
            with col_b:
                if st.button("Trim text spaces", use_container_width=True):
                    cleaned = current.copy()
                    for column in cleaned.select_dtypes(include=["object"]).columns:
                        cleaned[column] = cleaned[column].map(
                            lambda value: value.strip() if isinstance(value, str) else value
                        )
                    st.session_state.sheets[current_name] = cleaned
                    st.success("Text-এর অতিরিক্ত শুরু/শেষের space সরানো হয়েছে।")
                    st.rerun()
            with col_c:
                if st.button("Clear this worksheet", use_container_width=True):
                    st.session_state.sheets[current_name] = pd.DataFrame()
                    st.session_state.ai_result = ""
                    st.rerun()

            st.divider()
            st.subheader("⬇️ Export workbook")
            st.caption("সব worksheet একসঙ্গে একটি formatted XLSX ফাইলে export হবে।")
            try:
                export_sheets = OrderedDict(
                    (name, salary_calculations(frame))
                    for name, frame in st.session_state.sheets.items()
                )
                excel_bytes = make_excel(export_sheets)
                st.download_button(
                    "⬇️ Download Professional Excel Workbook",
                    data=excel_bytes,
                    file_name=f"AnalyzeData_{date.today().isoformat()}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            except Exception as error:
                st.error(f"Excel তৈরি করা যায়নি: {error}")

            st.divider()
            st.subheader("🤖 Gemini AI Data Assistant")
            question = st.text_area(
                "এই worksheet সম্পর্কে AI-কে কী জিজ্ঞাসা করবেন?",
                placeholder="এই data-এর summary দিন, missing value বা unusual record খুঁজুন...",
                key="gemini_question",
            )
            if st.button("Ask Gemini AI"):
                if question.strip():
                    with st.spinner("AI বিশ্লেষণ করছে..."):
                        try:
                            st.session_state.ai_result = gemini_answer(question, current)
                        except Exception as error:
                            st.session_state.ai_result = f"Gemini error: {error}"
                else:
                    st.warning("আগে একটি প্রশ্ন লিখুন।")
            if st.session_state.ai_result:
                st.markdown(st.session_state.ai_result)


# ------------------------------
# OFFICIAL TEMPLATES
# ------------------------------
elif page == "📁 Official Templates":
    st.title("📁 Official Work Templates")
    st.write(
        "বিভিন্ন ধরনের অফিসিয়াল কাজের জন্য ready-made column structures। "
        "Template বেছে নিয়ে নিজের data যোগ করতে পারবেন।"
    )
    category = st.selectbox("কাজের বিভাগ", list(TEMPLATES.keys()))
    template_names = list(TEMPLATES[category].keys())
    chosen = st.selectbox("Template নির্বাচন করুন", template_names)
    columns = TEMPLATES[category][chosen]
    st.subheader(chosen)
    st.caption(f"{category} · {len(columns)}টি column")
    st.dataframe(pd.DataFrame(columns=columns), use_container_width=True, hide_index=True)
    if st.button("Use this template", use_container_width=True):
        sheet_name = chosen
        st.session_state.sheets[sheet_name] = pd.DataFrame(columns=columns)
        st.session_state.active_sheet = sheet_name
        st.session_state.ai_result = ""
        st.success(
            f"'{chosen}' template তৈরি হয়েছে। Universal Data Workspace-এ গিয়ে data যোগ করুন।"
        )


# ------------------------------
# ABOUT
# ------------------------------
else:
    st.title("ℹ️ About AnalyzeData.io")
    st.write(
        "AnalyzeData.io হলো একটি configurable Universal Official Work ও Excel "
        "workspace। এটি বিভিন্ন ধরনের tabular office/business data import, inspect, "
        "edit, analyze, template তৈরি এবং Excel export করতে সাহায্য করে।"
    )
    st.markdown(
        """
        **বর্তমান সুবিধা**
        - CSV, XLSX, XLSM, XLS ও TSV import
        - Excel workbook-এর একাধিক worksheet import
        - প্রথম দিকের rows থেকে সম্ভাব্য heading শনাক্ত করার চেষ্টা
        - Data paste, edit ও নতুন table structure তৈরি
        - Missing value, duplicate row ও column profile
        - Numeric summary ও salary calculation
        - HR, Finance, Business, Education, Healthcare, Project, Logistics ইত্যাদির templates
        - Formatting ও heading-সহ multi-sheet Excel download
        - Gemini API key থাকলে AI data assistant

        **গুরুত্বপূর্ণ সীমাবদ্ধতা**
        - Heading detection স্বয়ংক্রিয় অনুমান; import-এর পরে Preview পরীক্ষা করা উচিত।
        - সব দেশের সব সরকারি/বেসরকারি workflow এক কোডে আগে থেকেই সম্পূর্ণ তৈরি থাকে না।
          Template ও configurable data workspace দিয়ে ধাপে ধাপে নতুন workflow যোগ করা যাবে।
        - Tax, payroll, medical, legal বা regulatory হিসাব ব্যবহারের আগে সংশ্লিষ্ট দেশের
          নিয়ম অনুযায়ী যাচাই করতে হবে।
        """
    )
    st.info(
        "Gemini AI চালু করতে Streamlit Cloud → App settings → Secrets-এ "
        "GEMINI_API_KEY যোগ করুন। API key কখনো app.py-তে লিখবেন না।"
    )