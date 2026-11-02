import io
import os
import re
import csv
from datetime import date

import pandas as pd
import streamlit as st
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

# -------------------- Page setup --------------------
st.set_page_config(
    page_title="AnalyzeData.io | AI Excel Generator",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Strong contrast: readable inputs, sidebar, tables and buttons.
st.markdown("""
<style>
:root { color-scheme: light; }
.stApp { background: #f4f7fc; color: #172b45; }
[data-testid="stSidebar"] { background: #102b4c; }
[data-testid="stSidebar"] * { color: #f4f8ff !important; }
[data-testid="stSidebar"] [data-testid="stRadio"] label {
    background: rgba(255,255,255,.07); border: 1px solid rgba(255,255,255,.13);
    border-radius: 10px; padding: .55rem .7rem; margin: .18rem 0;
}
[data-testid="stSidebar"] [data-testid="stRadio"] label:has(input:checked) {
    background: #2563eb; border-color: #78a7ff;
}
.hero {
    padding: 2rem; border-radius: 20px; color: white;
    background: linear-gradient(120deg,#102b4c,#2563eb);
    margin-bottom: 1rem;
}
.hero h1, .hero p { color: white !important; }
.metric-card {
    background: white; border: 1px solid #dce5f1; border-radius: 14px;
    padding: 1rem; min-height: 105px; box-shadow: 0 2px 8px rgba(16,43,76,.04);
}
.metric-label { color:#526783; font-size:.85rem; }
.metric-value { color:#102b4c; font-size:1.55rem; font-weight:750; }
h1,h2,h3,p,label { color:#172b45; }
.stTextArea textarea, .stTextInput input, .stNumberInput input {
    color:#172b45 !important; background:#ffffff !important;
    border:1px solid #aebed2 !important;
}
.stSelectbox div[data-baseweb="select"] { background:#fff; }
.stButton button, .stDownloadButton button {
    background:#2563eb; color:#fff; border:1px solid #2563eb;
    border-radius:10px; font-weight:650; min-height:2.65rem;
}
.stButton button:hover, .stDownloadButton button:hover {
    background:#1748b8; color:#fff; border-color:#1748b8;
}
div[data-testid="stDataFrame"], div[data-testid="stDataEditor"] {
    background:white; border-radius:12px;
}
.small-note { color:#526783; font-size:.9rem; }
</style>
""", unsafe_allow_html=True)

# -------------------- Helpers --------------------
def parse_pasted_data(raw: str) -> pd.DataFrame:
    """Parse CSV, TSV, pipe-separated, or whitespace-delimited pasted data."""
    raw = (raw or "").strip()
    if not raw:
        return pd.DataFrame()
    sample = "\n".join(raw.splitlines()[:8])
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",\t;|")
        df = pd.read_csv(io.StringIO(raw), sep=dialect.delimiter, engine="python")
        if len(df.columns) > 1:
            return clean_dataframe(df)
    except Exception:
        pass
    # Fall back to pandas' delimiter inference.
    try:
        df = pd.read_csv(io.StringIO(raw), sep=None, engine="python")
        if len(df.columns) > 1:
            return clean_dataframe(df)
    except Exception:
        pass
    lines = [re.split(r"\s{2,}", line.strip()) for line in raw.splitlines() if line.strip()]
    if len(lines) >= 2 and len({len(x) for x in lines}) == 1 and len(lines[0]) > 1:
        return clean_dataframe(pd.DataFrame(lines[1:], columns=lines[0]))
    return clean_dataframe(pd.DataFrame({"Data": raw.splitlines()}))

def clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() or f"Column_{i+1}" for i, c in enumerate(df.columns)]
    df = df.loc[:, ~pd.Index(df.columns).duplicated()]
    df = df.dropna(how="all")
    return df.reset_index(drop=True)

def read_upload(uploaded):
    name = uploaded.name.lower()
    data = uploaded.getvalue()
    if name.endswith((".xlsx", ".xlsm")):
        return clean_dataframe(pd.read_excel(io.BytesIO(data)))
    if name.endswith(".xls"):
        return clean_dataframe(pd.read_excel(io.BytesIO(data)))
    if name.endswith(".csv"):
        try:
            return clean_dataframe(pd.read_csv(io.BytesIO(data), sep=None, engine="python"))
        except UnicodeDecodeError:
            return clean_dataframe(pd.read_csv(io.BytesIO(data), encoding="latin-1", sep=None, engine="python"))
    if name.endswith(".tsv"):
        return clean_dataframe(pd.read_csv(io.BytesIO(data), sep="\t"))
    raise ValueError("এই ফাইল ফরম্যাট সমর্থিত নয়। CSV, XLSX বা XLS ব্যবহার করুন।")

def salary_calculations(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    normalized = {str(c).strip().lower(): c for c in out.columns}
    def find(*names):
        for n in names:
            if n in normalized:
                return normalized[n]
        return None
    basic = find("basic salary", "basic", "basic pay")
    hra = find("hra", "house rent allowance")
    allowance = find("other allowance", "allowance", "other allowances")
    deduction = find("deduction", "deductions")
    gross = find("gross salary", "gross")
    net = find("net salary", "net pay", "net")
    numeric_cols = [c for c in [basic, hra, allowance, deduction, gross, net] if c]
    for c in numeric_cols:
        out[c] = pd.to_numeric(out[c].astype(str).str.replace(",", "", regex=False), errors="coerce").fillna(0)
    if basic:
        if not hra:
            hra = "HRA"
            out[hra] = 0
        if not allowance:
            allowance = "Other Allowance"
            out[allowance] = 0
        if not gross:
            gross = "Gross Salary"
        out[gross] = out[basic] + out[hra] + out[allowance]
        if not deduction:
            deduction = "Deduction"
            out[deduction] = 0
        if not net:
            net = "Net Salary"
        out[net] = out[gross] - out[deduction]
    return out

def make_excel(df: pd.DataFrame, sheet_name="Data") -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name[:31])
        ws = writer.sheets[sheet_name[:31]]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        header_fill = PatternFill("solid", fgColor="173F73")
        header_font = Font(bold=True, color="FFFFFF")
        thin_blue = Side(style="thin", color="D8E3F0")
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = Border(bottom=thin_blue)
        for col_idx, col_cells in enumerate(ws.columns, 1):
            values = [len(str(c.value)) if c.value is not None else 0 for c in col_cells]
            width = min(max(max(values, default=10) + 3, 12), 34)
            ws.column_dimensions[get_column_letter(col_idx)].width = width
        ws.row_dimensions[1].height = 25
    return buffer.getvalue()

def get_gemini_client():
    api_key = st.secrets.get("GEMINI_API_KEY", None) if hasattr(st, "secrets") else None
    api_key = api_key or os.getenv("GEMINI_API_KEY")
    if not api_key:
        return None
    try:
        from google import genai
        return genai.Client(api_key=api_key)
    except Exception:
        return None

def gemini_text(prompt: str, df: pd.DataFrame | None = None) -> str:
    client = get_gemini_client()
    if client is None:
        return "Gemini সংযুক্ত নেই। Streamlit app-এর Secrets-এ GEMINI_API_KEY যোগ করুন।"
    context = ""
    if df is not None:
        context = "\n\nData sample (first 30 rows):\n" + df.head(30).to_csv(index=False)
    response = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=prompt + context
    )
    return response.text or "AI কোনো লেখা ফেরত দেয়নি।"

# -------------------- Session state --------------------
if "df" not in st.session_state:
    st.session_state.df = pd.DataFrame()
if "ai_result" not in st.session_state:
    st.session_state.ai_result = ""

# -------------------- Sidebar navigation --------------------
with st.sidebar:
    st.markdown("## 📊 AnalyzeData.io")
    st.caption("AI Excel Generator")
    st.markdown("---")
    page = st.radio(
        "NAVIGATION",
        ["🏠 Dashboard", "✨ AI Excel Generator", "📁 Templates", "ℹ️ About"],
        label_visibility="visible"
    )
    st.markdown("---")
    st.markdown("### CURRENT FEATURES")
    st.markdown("🟢 **CSV / XLSX Import**")
    st.markdown("🟢 **Paste & edit data**")
    st.markdown("🟢 **Excel export**")
    st.markdown("🟢 **Salary calculations**")
    st.markdown("🟡 **Gemini AI** — API key required")
    st.markdown("---")
    st.caption("AnalyzeData.io · Built with Streamlit")

# -------------------- Dashboard --------------------
if page == "🏠 Dashboard":
    st.markdown("""
    <div class="hero">
      <h1>AnalyzeData.io</h1>
      <p>Turn messy data into clean, useful Excel workbooks.</p>
      <p>Paste data · Upload files · Edit preview · Analyze with AI · Download Excel</p>
    </div>
    """, unsafe_allow_html=True)
    a, b, c = st.columns(3)
    with a:
        st.markdown('<div class="metric-card"><div class="metric-label">Workbook builder</div><div class="metric-value">Ready</div></div>', unsafe_allow_html=True)
    with b:
        st.markdown('<div class="metric-card"><div class="metric-label">Editable preview</div><div class="metric-value">Included</div></div>', unsafe_allow_html=True)
    with c:
        st.markdown('<div class="metric-card"><div class="metric-label">Excel export</div><div class="metric-value">Included</div></div>', unsafe_allow_html=True)
    st.subheader("শুরু করুন")
    st.write("বাম পাশের **AI Excel Generator** থেকে ডেটা পেস্ট করুন অথবা CSV/XLSX ফাইল আপলোড করুন।")
    if not st.session_state.df.empty:
        st.success(f"বর্তমান ডেটায় {len(st.session_state.df)}টি row এবং {len(st.session_state.df.columns)}টি column আছে।")
        st.dataframe(st.session_state.df.head(8), use_container_width=True)

# -------------------- Main generator --------------------
elif page == "✨ AI Excel Generator":
    st.title("✨ AI Excel Generator")
    st.caption("ডেটা পেস্ট করুন বা ফাইল আপলোড করুন। Preview সম্পাদনা করে Excel ডাউনলোড করুন।")
    tab_paste, tab_upload = st.tabs(["📝 Paste data", "📤 Upload file"])
    with tab_paste:
        raw = st.text_area(
            "টেবিল ডেটা পেস্ট করুন (প্রথম লাইনে column heading রাখুন)",
            height=180,
            placeholder="Employee ID,Employee Name,Department,Basic Salary,HRA,Other Allowance,Deduction\nEMP101,Rahul Das,Sales,18000,3000,2000,500"
        )
        if st.button("Convert pasted data", use_container_width=True):
            parsed = parse_pasted_data(raw)
            if parsed.empty:
                st.warning("আগে কিছু ডেটা পেস্ট করুন।")
            else:
                st.session_state.df = parsed
                st.success(f"{len(parsed)}টি row এবং {len(parsed.columns)}টি column পাওয়া গেছে।")
    with tab_upload:
        uploaded = st.file_uploader("CSV, XLSX বা XLS ফাইল বেছে নিন", type=["csv", "xlsx", "xls", "tsv"])
        if uploaded is not None and st.button("Import uploaded file", use_container_width=True):
            try:
                st.session_state.df = read_upload(uploaded)
                st.success(f"{len(st.session_state.df)}টি row import হয়েছে।")
            except Exception as e:
                st.error(f"ফাইল পড়া যায়নি: {e}")

    st.divider()
    st.subheader("নতুন টেবিল তৈরি করুন")
    description = st.text_input("কোন কোন column চান?", placeholder="Name, Department, Phone, Joining Date")
    if st.button("Create empty structure"):
        cols = [x.strip() for x in re.split(r"[,;|]", description) if x.strip()]
        if cols:
            st.session_state.df = pd.DataFrame(columns=cols)
            st.success("নতুন empty structure তৈরি হয়েছে।")
        else:
            st.warning("কমা দিয়ে column-এর নাম লিখুন।")

    if not st.session_state.df.empty:
        st.divider()
        st.subheader("📊 Excel Preview & Editor")
        st.caption("নিচের cell-এ ক্লিক করে ডেটা পরিবর্তন করতে পারবেন।")
        edited = st.data_editor(
            st.session_state.df,
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True,
            key="main_data_editor"
        )
        st.session_state.df = edited
        left, right = st.columns(2)
        with left:
            if st.button("🧮 Calculate salary", use_container_width=True):
                st.session_state.df = salary_calculations(st.session_state.df)
                st.success("Salary calculation সম্পন্ন হয়েছে।")
                st.rerun()
        with right:
            if st.button("🗑️ Clear current data", use_container_width=True):
                st.session_state.df = pd.DataFrame()
                st.rerun()
        excel_bytes = make_excel(st.session_state.df)
        st.download_button(
            "⬇️ Download Professional Excel",
            data=excel_bytes,
            file_name=f"AnalyzeData_{date.today().isoformat()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

        st.divider()
        st.subheader("🤖 Gemini AI Data Assistant")
        st.caption("Gemini API key যুক্ত থাকলে AI আপনার ডেটা বিশ্লেষণ করতে পারবে।")
        question = st.text_area("আপনার প্রশ্ন", placeholder="এই ডেটার summary দিন, ভুল বা missing value খুঁজুন...")
        if st.button("Ask Gemini AI", use_container_width=True):
            if question.strip():
                with st.spinner("AI বিশ্লেষণ করছে..."):
                    try:
                        st.session_state.ai_result = gemini_text(question, st.session_state.df)
                    except Exception as e:
                        st.session_state.ai_result = f"Gemini error: {e}"
            else:
                st.warning("আগে একটি প্রশ্ন লিখুন।")
        if st.session_state.ai_result:
            st.markdown(st.session_state.ai_result)

# -------------------- Templates --------------------
elif page == "📁 Templates":
    st.title("📁 Ready-made templates")
    st.write("একটি template বেছে নিয়ে নিজের ডেটা দিয়ে পূরণ করুন।")
    templates = {
        "Employee Salary": ["Employee ID", "Employee Name", "Department", "Basic Salary", "HRA", "Other Allowance", "Deduction"],
        "Student Marks": ["Roll", "Student Name", "Subject", "Total Marks", "Marks Obtained", "Grade"],
        "Sales Report": ["Date", "Product", "Category", "Quantity", "Unit Price", "Total Sales"],
        "Attendance": ["Date", "Employee Name", "Department", "Status", "In Time", "Out Time"],
        "Expense Tracker": ["Date", "Description", "Category", "Payment Method", "Amount"],
    }
    choice = st.selectbox("Template নির্বাচন করুন", list(templates))
    st.dataframe(pd.DataFrame(columns=templates[choice]), use_container_width=True)
    if st.button("Use this template"):
        st.session_state.df = pd.DataFrame(columns=templates[choice])
        st.success("Template তৈরি হয়েছে। AI Excel Generator পেজে যান।")

# -------------------- About --------------------
else:
    st.title("ℹ️ About AnalyzeData.io")
    st.write("AnalyzeData.io একটি data-to-Excel workbook builder।")
    st.markdown("""
    **যা করতে পারবেন**
    - CSV, XLSX এবং TSV import
    - CSV/TSV/pipe-separated data paste
    - Table preview ও cell editing
    - Basic salary, HRA, allowance, gross, deduction ও net salary হিসাব
    - Formatted Excel workbook download
    - Gemini API key থাকলে AI data analysis
    """)
    st.info("AI ব্যবহার করতে Streamlit Cloud → App settings → Secrets-এ GEMINI_API_KEY যোগ করুন। API key কখনো app.py-তে লিখবেন না।")