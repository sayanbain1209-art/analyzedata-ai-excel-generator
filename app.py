import io
import os
import re
import csv
from datetime import date

import pandas as pd
import streamlit as st
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo


# ---------------------------------------------------------
# PAGE CONFIGURATION
# ---------------------------------------------------------
st.set_page_config(
    page_title="AnalyzeData.io | AI Excel Generator",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    :root { color-scheme: light; }

    .stApp {
        background: #f4f7fc;
        color: #172b45;
    }

    [data-testid="stSidebar"] {
        background: #102b4c;
    }

    [data-testid="stSidebar"] * {
        color: #f4f8ff !important;
    }

    [data-testid="stSidebar"] [data-testid="stRadio"] label {
        background: rgba(255,255,255,.08);
        border: 1px solid rgba(255,255,255,.18);
        border-radius: 10px;
        padding: .55rem .7rem;
        margin: .18rem 0;
    }

    [data-testid="stSidebar"] [data-testid="stRadio"] label:has(input:checked) {
        background: #2563eb;
        border-color: #9bc0ff;
    }

    .hero {
        padding: 2rem;
        border-radius: 20px;
        color: white;
        background: linear-gradient(120deg, #102b4c, #2563eb);
        margin-bottom: 1rem;
    }

    .hero h1, .hero p {
        color: white !important;
    }

    .metric-card {
        background: white;
        border: 1px solid #dce5f1;
        border-radius: 14px;
        padding: 1rem;
        min-height: 95px;
    }

    .metric-label {
        color: #526783;
        font-size: .85rem;
    }

    .metric-value {
        color: #102b4c;
        font-size: 1.5rem;
        font-weight: 750;
    }

    h1, h2, h3, p, label {
        color: #172b45;
    }

    .stTextArea textarea, .stTextInput input {
        color: #172b45 !important;
        background: white !important;
    }

    .stButton button, .stDownloadButton button {
        background: #2563eb;
        color: white;
        border: 1px solid #2563eb;
        border-radius: 10px;
        font-weight: 650;
        min-height: 2.65rem;
    }

    .stButton button:hover, .stDownloadButton button:hover {
        background: #1748b8;
        color: white;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# ---------------------------------------------------------
# DATA CLEANING AND IMPORT
# ---------------------------------------------------------
def clean_df(df):
    """Clean column names and remove completely empty rows/duplicate columns."""
    df = df.copy()

    cleaned_columns = []
    for i, column in enumerate(df.columns):
        name = str(column).strip()
        if not name or name.lower().startswith("unnamed:"):
            name = f"Column_{i + 1}"
        cleaned_columns.append(name)

    df.columns = cleaned_columns
    df = df.loc[:, ~pd.Index(df.columns).duplicated()]
    df = df.dropna(how="all")

    return df.reset_index(drop=True)


def parse_data(raw):
    """Read comma, tab, semicolon or pipe-separated pasted data."""
    raw = (raw or "").strip()

    if not raw:
        return pd.DataFrame()

    sample = "\n".join(raw.splitlines()[:8])

    try:
        dialect = csv.Sniffer().sniff(
            sample,
            delimiters=",\t;|",
        )
        df = pd.read_csv(
            io.StringIO(raw),
            sep=dialect.delimiter,
            engine="python",
        )
        if len(df.columns) > 1:
            return clean_df(df)
    except Exception:
        pass

    try:
        df = pd.read_csv(
            io.StringIO(raw),
            sep=None,
            engine="python",
        )
        if len(df.columns) > 1:
            return clean_df(df)
    except Exception:
        pass

    raise ValueError(
        "প্রথম লাইনে column heading দিন এবং প্রতিটি column "
        "কমা অথবা Tab দিয়ে আলাদা করুন।"
    )


# ---------------------------------------------------------
# SALARY CALCULATIONS
# ---------------------------------------------------------
def salary_calculations(df):
    out = clean_df(df)
    columns = {str(c).strip().lower(): c for c in out.columns}

    def find(*names):
        for name in names:
            if name in columns:
                return columns[name]
        return None

    basic = find("basic salary", "basic", "basic pay")

    if basic is None:
        return out

    hra = find("hra", "house rent allowance")
    allowance = find(
        "other allowance",
        "allowance",
        "other allowances",
    )
    deduction = find("deduction", "deductions")

    def numbers(column):
        if column is None:
            return pd.Series(0, index=out.index, dtype="float64")

        return pd.to_numeric(
            out[column]
            .astype(str)
            .str.replace(",", "", regex=False)
            .str.strip(),
            errors="coerce",
        ).fillna(0)

    out["Gross Salary"] = (
        numbers(basic) + numbers(hra) + numbers(allowance)
    )
    out["Net Salary"] = out["Gross Salary"] - numbers(deduction)

    return out


# ---------------------------------------------------------
# EXCEL EXPORT
# ---------------------------------------------------------
def make_excel(df):
    """Create an Excel file with a real Excel Table and visible row-1 headers."""
    df = clean_df(df)

    if df.empty or len(df.columns) == 0:
        raise ValueError("Excel export করার মতো data নেই।")

    # Keep all real columns; clean_df has already repaired blank/unnamed names.
    headers = [str(column).strip() for column in df.columns]
    if any(not heading for heading in headers):
        raise ValueError("এক বা একাধিক column heading ফাঁকা আছে।")

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Data"
    worksheet.sheet_state = "visible"
    worksheet.sheet_view.showGridLines = True
    worksheet.sheet_view.showRowColHeaders = True
    worksheet.sheet_view.zoomScale = 90

    # Write the header row explicitly, then write records starting at row 2.
    for col_num, heading in enumerate(headers, start=1):
        cell = worksheet.cell(row=1, column=col_num, value=heading)
        cell.font = Font(name="Arial", bold=True, color="FFFFFF", size=11)
        cell.fill = PatternFill(fill_type="solid", fgColor="1F4E78")
        cell.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=False
        )
        cell.border = Border(
            top=Side(style="thin", color="1F1F1F"),
            bottom=Side(style="medium", color="1F1F1F"),
            left=Side(style="thin", color="D9E2F3"),
            right=Side(style="thin", color="D9E2F3"),
        )

    for row_num, record in enumerate(df.itertuples(index=False, name=None), start=2):
        for col_num, value in enumerate(record, start=1):
            if pd.isna(value):
                value = None
            worksheet.cell(row=row_num, column=col_num, value=value)

    worksheet.row_dimensions[1].height = 28
    # Do not freeze panes: some mobile spreadsheet viewers render frozen
    # first rows incorrectly. The Excel Table below supplies filter controls.
    worksheet.freeze_panes = None

    # Set readable widths, with enough room for the full heading text.
    for col_num, heading in enumerate(headers, start=1):
        letter = get_column_letter(col_num)
        longest_data = max(
            (len(str(df.iloc[row, col_num - 1])) for row in range(len(df))),
            default=0,
        )
        width = max(len(heading) + 4, min(longest_data + 3, 30), 16)
        worksheet.column_dimensions[letter].width = min(width, 34)

    for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row):
        for cell in row:
            cell.alignment = Alignment(vertical="center", wrap_text=False)
            cell.border = Border(
                bottom=Side(style="thin", color="D8E3F0")
            )

    # A genuine Excel Table marks row 1 as the header row in Excel/mobile apps.
    table = Table(displayName="AnalyzeDataTable", ref=worksheet.dimensions)
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    worksheet.add_table(table)

    output = io.BytesIO()
    workbook.save(output)
    excel_bytes = output.getvalue()

    # Re-open the actual saved bytes and verify every header is in row 1.
    check_workbook = load_workbook(io.BytesIO(excel_bytes), read_only=False, data_only=False)
    check_sheet = check_workbook["Data"]
    saved_headers = [
        check_sheet.cell(row=1, column=col_num).value
        for col_num in range(1, len(headers) + 1)
    ]
    has_table = "AnalyzeDataTable" in check_sheet.tables
    check_workbook.close()

    if saved_headers != headers or not has_table:
        raise ValueError("Excel export validation failed: header row/table is missing.")

    return excel_bytes


# ---------------------------------------------------------
# GEMINI AI
# ---------------------------------------------------------
def gemini_answer(prompt, df):
    try:
        key = st.secrets.get("GEMINI_API_KEY", None)
    except Exception:
        key = None

    key = key or os.getenv("GEMINI_API_KEY")

    if not key:
        return (
            "Gemini এখনো সংযুক্ত নয়। Streamlit Cloud → App settings "
            "→ Secrets-এ GEMINI_API_KEY যোগ করুন।"
        )

    from google import genai

    client = genai.Client(api_key=key)

    context = (
        "\n\nData sample (up to 30 rows):\n"
        + df.head(30).to_csv(index=False)
    )

    result = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=prompt + context,
    )

    return result.text or "AI কোনো উত্তর ফেরত দেয়নি।"


# ---------------------------------------------------------
# SESSION STATE
# ---------------------------------------------------------
if "df" not in st.session_state:
    st.session_state.df = pd.DataFrame()

if "ai_result" not in st.session_state:
    st.session_state.ai_result = ""


# ---------------------------------------------------------
# SIDEBAR NAVIGATION
# ---------------------------------------------------------
with st.sidebar:
    st.markdown("## 📊 AnalyzeData.io")
    st.caption("AI Excel Generator")
    st.markdown("---")

    page = st.radio(
        "NAVIGATION",
        [
            "🏠 Dashboard",
            "✨ AI Excel Generator",
            "📁 Templates",
            "ℹ️ About",
        ],
    )

    st.markdown("---")
    st.markdown("### CURRENT FEATURES")
    st.markdown(
        """
        🟢 **CSV / XLSX Import**

        🟢 **Paste & edit data**

        🟢 **Excel export with headings**

        🟢 **Salary calculations**

        🟡 **Gemini AI** — API key required
        """
    )


# ---------------------------------------------------------
# DASHBOARD
# ---------------------------------------------------------
if page == "🏠 Dashboard":
    st.markdown(
        """
        <div class="hero">
            <h1>AnalyzeData.io</h1>
            <p>Turn messy data into clean, useful Excel workbooks.</p>
            <p>Paste data · Upload files · Edit preview · Analyze with AI · Download Excel</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    col1, col2, col3 = st.columns(3)

    cards = [
        (col1, "Workbook builder", "Ready"),
        (col2, "Editable preview", "Included"),
        (col3, "Excel export", "Included"),
    ]

    for column, label, value in cards:
        with column:
            st.markdown(
                f"""
                <div class="metric-card">
                    <div class="metric-label">{label}</div>
                    <div class="metric-value">{value}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

    st.subheader("শুরু করুন")
    st.write(
        "বাম পাশের **AI Excel Generator** থেকে ডেটা পেস্ট করুন "
        "অথবা CSV/XLSX ফাইল আপলোড করুন।"
    )

    if not st.session_state.df.empty:
        st.dataframe(
            st.session_state.df,
            use_container_width=True,
        )


# ---------------------------------------------------------
# AI EXCEL GENERATOR
# ---------------------------------------------------------
elif page == "✨ AI Excel Generator":
    st.title("✨ AI Excel Generator")
    st.caption(
        "ডেটা পেস্ট করুন বা ফাইল আপলোড করুন। "
        "Preview সম্পাদনা করে Excel ডাউনলোড করুন।"
    )

    paste_tab, upload_tab = st.tabs(
        ["📝 Paste data", "📤 Upload file"]
    )

    with paste_tab:
        raw = st.text_area(
            "টেবিল ডেটা পেস্ট করুন (প্রথম লাইনে column heading রাখুন)",
            height=180,
            placeholder=(
                "Employee ID,Employee Name,Department,Basic Salary,"
                "HRA,Other Allowance,Deduction\n"
                "EMP101,Rahul Das,Sales,18000,3000,2000,500"
            ),
        )

        if st.button(
            "Convert pasted data",
            use_container_width=True,
        ):
            try:
                parsed = parse_data(raw)

                if parsed.empty:
                    st.warning("আগে কিছু ডেটা পেস্ট করুন।")
                else:
                    st.session_state.df = parsed
                    st.session_state.ai_result = ""
                    st.success(
                        f"{len(parsed)}টি row এবং "
                        f"{len(parsed.columns)}টি column পাওয়া গেছে।"
                    )
            except Exception as error:
                st.error(str(error))

    with upload_tab:
        uploaded_file = st.file_uploader(
            "CSV, XLSX বা XLS ফাইল বেছে নিন",
            type=["csv", "xlsx", "xls", "tsv"],
        )

        if uploaded_file is not None:
            if st.button(
                "Import uploaded file",
                use_container_width=True,
            ):
                try:
                    file_data = uploaded_file.getvalue()
                    file_name = uploaded_file.name.lower()

                    if file_name.endswith((".xlsx", ".xlsm", ".xls")):
                        imported_df = pd.read_excel(
                            io.BytesIO(file_data)
                        )
                    elif file_name.endswith(".tsv"):
                        imported_df = pd.read_csv(
                            io.BytesIO(file_data),
                            sep="\t",
                        )
                    else:
                        imported_df = pd.read_csv(
                            io.BytesIO(file_data),
                            sep=None,
                            engine="python",
                        )

                    st.session_state.df = clean_df(imported_df)
                    st.session_state.ai_result = ""

                    st.success(
                        f"{len(st.session_state.df)}টি row import হয়েছে।"
                    )
                except Exception as error:
                    st.error(f"ফাইল পড়া যায়নি: {error}")

    st.divider()
    st.subheader("নতুন টেবিল তৈরি করুন")

    description = st.text_input(
        "কোন কোন column চান?",
        placeholder="Name, Department, Phone, Joining Date",
    )

    if st.button("Create empty structure"):
        columns = [
            item.strip()
            for item in re.split(r"[,;|]", description)
            if item.strip()
        ]

        if columns:
            st.session_state.df = pd.DataFrame(columns=columns)
            st.session_state.ai_result = ""
            st.success("নতুন empty structure তৈরি হয়েছে।")
        else:
            st.warning("কমা দিয়ে column-এর নাম লিখুন।")

    # Show editor and export tools only when data exists.
    if not st.session_state.df.empty:
        st.divider()
        st.subheader("📊 Excel Preview & Editor")
        st.caption(
            "নিচের cell-এ ক্লিক করে ডেটা পরিবর্তন করতে পারবেন। "
            "Column heading-ও সম্পাদনা করা যাবে।"
        )

        edited_df = st.data_editor(
            st.session_state.df,
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True,
            key="main_data_editor",
        )

        st.session_state.df = clean_df(edited_df)

        left, right = st.columns(2)

        with left:
            if st.button(
                "🧮 Calculate salary",
                use_container_width=True,
            ):
                old_df = st.session_state.df.copy()
                new_df = salary_calculations(old_df)

                if new_df.equals(old_df):
                    st.warning(
                        "Basic Salary column পাওয়া যায়নি। "
                        "Column heading পরীক্ষা করুন।"
                    )
                else:
                    st.session_state.df = new_df
                    st.success(
                        "Gross Salary এবং Net Salary হিসাব করা হয়েছে।"
                    )
                    st.rerun()

        with right:
            if st.button(
                "🗑️ Clear current data",
                use_container_width=True,
            ):
                st.session_state.df = pd.DataFrame()
                st.session_state.ai_result = ""
                st.rerun()

        st.subheader("Excel download")
        st.caption(
            "Excel ফাইলে প্রথম row-তে সব column heading থাকবে। "
            "Employee data শুরু হবে দ্বিতীয় row থেকে।"
        )

        try:
            export_df = salary_calculations(
                st.session_state.df
            )

            excel_bytes = make_excel(export_df)

            st.download_button(
                "⬇️ Download Professional Excel",
                data=excel_bytes,
                file_name=f"AnalyzeData_{date.today().isoformat()}.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
                use_container_width=True,
            )
        except Exception as error:
            st.error(f"Excel তৈরি করা যায়নি: {error}")

        st.divider()
        st.subheader("🤖 Gemini AI Data Assistant")

        question = st.text_area(
            "আপনার প্রশ্ন",
            placeholder=(
                "এই ডেটার summary দিন, ভুল বা missing value খুঁজুন..."
            ),
        )

        if st.button("Ask Gemini AI"):
            if question.strip():
                with st.spinner("AI বিশ্লেষণ করছে..."):
                    try:
                        st.session_state.ai_result = gemini_answer(
                            question,
                            st.session_state.df,
                        )
                    except Exception as error:
                        st.session_state.ai_result = (
                            f"Gemini error: {error}"
                        )
            else:
                st.warning("আগে একটি প্রশ্ন লিখুন।")

        if st.session_state.ai_result:
            st.markdown(st.session_state.ai_result)


# ---------------------------------------------------------
# TEMPLATES
# ---------------------------------------------------------
elif page == "📁 Templates":
    st.title("📁 Ready-made templates")
    st.write(
        "একটি template বেছে নিয়ে নিজের ডেটা দিয়ে পূরণ করুন।"
    )

    templates = {
        "Employee Salary": [
            "Employee ID",
            "Employee Name",
            "Department",
            "Basic Salary",
            "HRA",
            "Other Allowance",
            "Deduction",
        ],
        "Student Marks": [
            "Roll",
            "Student Name",
            "Subject",
            "Total Marks",
            "Marks Obtained",
            "Grade",
        ],
        "Sales Report": [
            "Date",
            "Product",
            "Category",
            "Quantity",
            "Unit Price",
            "Total Sales",
        ],
        "Attendance": [
            "Date",
            "Employee Name",
            "Department",
            "Status",
            "In Time",
            "Out Time",
        ],
        "Expense Tracker": [
            "Date",
            "Description",
            "Category",
            "Payment Method",
            "Amount",
        ],
    }

    choice = st.selectbox(
        "Template নির্বাচন করুন",
        list(templates),
    )

    st.dataframe(
        pd.DataFrame(columns=templates[choice]),
        use_container_width=True,
    )

    if st.button("Use this template"):
        st.session_state.df = pd.DataFrame(
            columns=templates[choice]
        )
        st.session_state.ai_result = ""
        st.success(
            "Template তৈরি হয়েছে। AI Excel Generator পেজে যান।"
        )


# ---------------------------------------------------------
# ABOUT
# ---------------------------------------------------------
else:
    st.title("ℹ️ About AnalyzeData.io")
    st.write(
        "AnalyzeData.io একটি data-to-Excel workbook builder।"
    )

    st.markdown(
        """
        **যা করতে পারবেন**

        - CSV, XLSX এবং TSV import
        - Data paste ও cell editing
        - Salary calculation
        - Heading, filter ও formatting-সহ Excel download
        - Gemini API key থাকলে AI data analysis
        """
    )

    st.info(
        "AI ব্যবহার করতে Streamlit Cloud → App settings → Secrets-এ "
        "GEMINI_API_KEY যোগ করুন। API key কখনো app.py-তে লিখবেন না।"
    )
