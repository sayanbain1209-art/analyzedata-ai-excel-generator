import io, os, re, csv
from datetime import date
import pandas as pd
import streamlit as st
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(page_title='AnalyzeData.io | AI Excel Generator', page_icon='📊', layout='wide', initial_sidebar_state='expanded')
st.markdown('''<style>
:root{color-scheme:light}.stApp{background:#f4f7fc;color:#172b45}
[data-testid="stSidebar"]{background:#102b4c}[data-testid="stSidebar"] *{color:#f4f8ff!important}
[data-testid="stSidebar"] [data-testid="stRadio"] label{background:rgba(255,255,255,.08);border:1px solid rgba(255,255,255,.18);border-radius:10px;padding:.55rem .7rem;margin:.18rem 0}
[data-testid="stSidebar"] [data-testid="stRadio"] label:has(input:checked){background:#2563eb;border-color:#9bc0ff}
.hero{padding:2rem;border-radius:20px;color:white;background:linear-gradient(120deg,#102b4c,#2563eb);margin-bottom:1rem}.hero h1,.hero p{color:white!important}
.metric-card{background:white;border:1px solid #dce5f1;border-radius:14px;padding:1rem;min-height:95px}.metric-label{color:#526783;font-size:.85rem}.metric-value{color:#102b4c;font-size:1.5rem;font-weight:750}
h1,h2,h3,p,label{color:#172b45}.stTextArea textarea,.stTextInput input{color:#172b45!important;background:white!important}
.stButton button,.stDownloadButton button{background:#2563eb;color:white;border:1px solid #2563eb;border-radius:10px;font-weight:650;min-height:2.65rem}
.stButton button:hover,.stDownloadButton button:hover{background:#1748b8;color:white}
</style>''', unsafe_allow_html=True)

def clean_df(df):
    df=df.copy(); df.columns=[str(c).strip() if str(c).strip() else f'Column_{i+1}' for i,c in enumerate(df.columns)]
    df=df.loc[:,~pd.Index(df.columns).duplicated()].dropna(how='all')
    return df.reset_index(drop=True)

def parse_data(raw):
    raw=(raw or '').strip()
    if not raw: return pd.DataFrame()
    try:
        dialect=csv.Sniffer().sniff('\n'.join(raw.splitlines()[:8]),delimiters=',\t;|')
        df=pd.read_csv(io.StringIO(raw),sep=dialect.delimiter,engine='python')
        if len(df.columns)>1: return clean_df(df)
    except Exception: pass
    try:
        df=pd.read_csv(io.StringIO(raw),sep=None,engine='python')
        if len(df.columns)>1: return clean_df(df)
    except Exception: pass
    raise ValueError('প্রথম লাইনে column heading দিন এবং comma বা tab দিয়ে আলাদা করুন।')

def salary_calculations(df):
    out=clean_df(df); cols={str(c).strip().lower():c for c in out.columns}
    def find(*names):
        return next((cols[n] for n in names if n in cols),None)
    basic=find('basic salary','basic','basic pay')
    if basic is None: return out
    hra=find('hra','house rent allowance'); allowance=find('other allowance','allowance','other allowances'); deduction=find('deduction','deductions')
    def nums(col):
        if col is None: return pd.Series(0,index=out.index,dtype='float64')
        return pd.to_numeric(out[col].astype(str).str.replace(',','',regex=False).str.strip(),errors='coerce').fillna(0)
    out['Gross Salary']=nums(basic)+nums(hra)+nums(allowance)
    out['Net Salary']=out['Gross Salary']-nums(deduction)
    return out

def make_excel(df):
    df=clean_df(df)
    df=df.loc[:,[c for c in df.columns if str(c).strip() and not str(c).lower().startswith('unnamed:')]]
    if df.empty or len(df.columns)==0: raise ValueError('Excel export করার মতো data নেই।')
    bio=io.BytesIO()
    with pd.ExcelWriter(bio,engine='openpyxl') as writer:
        df.to_excel(writer,index=False,header=True,startrow=0,sheet_name='Data')
        ws=writer.sheets['Data']
        # Explicitly write headings into Excel row 1.
        for j,name in enumerate(df.columns,1): ws.cell(row=1,column=j,value=str(name))
        ws.freeze_panes='A2'; ws.auto_filter.ref=ws.dimensions
        fill=PatternFill('solid',fgColor='173F73'); font=Font(bold=True,color='FFFFFF')
        edge=Side(style='thin',color='D8E3F0')
        for cell in ws[1]:
            cell.fill=fill; cell.font=font; cell.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True); cell.border=Border(bottom=edge)
        for j,cells in enumerate(ws.columns,1):
            width=min(max(max((len(str(c.value)) if c.value is not None else 0) for c in cells)+3,12),34)
            ws.column_dimensions[get_column_letter(j)].width=width
        ws.row_dimensions[1].height=28
    bio.seek(0); return bio.getvalue()

def gemini_answer(prompt,df):
    key=st.secrets.get('GEMINI_API_KEY',None) if hasattr(st,'secrets') else None
    key=key or os.getenv('GEMINI_API_KEY')
    if not key: return 'Gemini সংযুক্ত নেই। App settings → Secrets-এ GEMINI_API_KEY যোগ করুন।'
    from google import genai
    client=genai.Client(api_key=key)
    context='\n\nData sample:\n'+df.head(30).to_csv(index=False)
    result=client.models.generate_content(model='gemini-2.5-flash',contents=prompt+context)
    return result.text or 'AI কোনো লেখা ফেরত দেয়নি।'

if 'df' not in st.session_state: st.session_state.df=pd.DataFrame()
if 'ai_result' not in st.session_state: st.session_state.ai_result=''
with st.sidebar:
    st.markdown('## 📊 AnalyzeData.io'); st.caption('AI Excel Generator'); st.markdown('---')
    page=st.radio('NAVIGATION',['🏠 Dashboard','✨ AI Excel Generator','📁 Templates','ℹ️ About'])
    st.markdown('---'); st.markdown('### CURRENT FEATURES')
    st.markdown('🟢 **CSV / XLSX Import**\n\n🟢 **Paste & edit data**\n\n🟢 **Excel export with headings**\n\n🟢 **Salary calculations**\n\n🟡 **Gemini AI** — API key required')

if page=='🏠 Dashboard':
    st.markdown('<div class="hero"><h1>AnalyzeData.io</h1><p>Turn messy data into clean, useful Excel workbooks.</p><p>Paste data · Upload files · Edit preview · Analyze with AI · Download Excel</p></div>',unsafe_allow_html=True)
    a,b,c=st.columns(3)
    for col,label,value in [(a,'Workbook builder','Ready'),(b,'Editable preview','Included'),(c,'Excel export','Included')]:
        with col: st.markdown(f'<div class="metric-card"><div class="metric-label">{label}</div><div class="metric-value">{value}</div></div>',unsafe_allow_html=True)
    st.subheader('শুরু করুন'); st.write('বাম পাশের **AI Excel Generator** থেকে ডেটা পেস্ট করুন অথবা CSV/XLSX ফাইল আপলোড করুন।')
    if not st.session_state.df.empty: st.dataframe(st.session_state.df,use_container_width=True)
elif page=='✨ AI Excel Generator':
    st.title('✨ AI Excel Generator'); st.caption('ডেটা পেস্ট করুন বা ফাইল আপলোড করুন। Preview সম্পাদনা করে Excel ডাউনলোড করুন।')
    paste,upload=st.tabs(['📝 Paste data','📤 Upload file'])
    with paste:
        raw=st.text_area('টেবিল ডেটা পেস্ট করুন (প্রথম লাইনে column heading রাখুন)',height=180,placeholder='Employee ID,Employee Name,Department,Basic Salary,HRA,Other Allowance,Deduction\nEMP101,Rahul Das,Sales,18000,3000,2000,500')
        if st.button('Convert pasted data',use_container_width=True):
            try:
                parsed=parse_data(raw)
                if parsed.empty: st.warning('আগে কিছু ডেটা পেস্ট করুন।')
                else: st.session_state.df=parsed; st.success(f'{len(parsed)}টি row এবং {len(parsed.columns)}টি column পাওয়া গেছে।')
            except Exception as e: st.error(str(e))
    with upload:
        file=st.file_uploader('CSV, XLSX বা XLS ফাইল বেছে নিন',type=['csv','xlsx','xls','tsv'])
        if file is not None and st.button('Import uploaded file',use_container_width=True):
            try:
                data=file.getvalue(); name=file.name.lower()
                if name.endswith(('.xlsx','.xlsm','.xls')): df=pd.read_excel(io.BytesIO(data))
                elif name.endswith('.tsv'): df=pd.read_csv(io.BytesIO(data),sep='\t')
                else: df=pd.read_csv(io.BytesIO(data),sep=None,engine='python')
                st.session_state.df=clean_df(df); st.success(f'{len(df)}টি row import হয়েছে।')
            except Exception as e: st.error(f'ফাইল পড়া যায়নি: {e}')
    st.divider(); st.subheader('নতুন টেবিল তৈরি করুন')
    desc=st.text_input('কোন কোন column চান?',placeholder='Name, Department, Phone, Joining Date')
    if st.button('Create empty structure'):
        columns=[x.strip() for x in re.split('[,;|]',desc) if x.strip()]
        if columns: st.session_state.df=pd.DataFrame(columns=columns); st.success('নতুন empty structure তৈরি হয়েছে।')
        else: st.warning('কমা দিয়ে column-এর নাম লিখুন।')
    if not st.session_state.df.empty:
        st.divider(); st.subheader('📊 Excel Preview & Editor'); st.caption('নিচের cell-এ ক্লিক করে ডেটা পরিবর্তন করতে পারবেন।')
        st.session_state.df=clean_df(st.data_editor(st.session_state.df,num_rows='dynamic',use_container_width=True,hide_index=True,key='main_data_editor'))
        left,right=st.columns(2)
        with left:
            if st.button('🧮 Calculate salary',use_container_width=True):
                old=st.session_state.df.copy(); new=salary_calculations(old)
                if new.equals(old): st.warning('Basic Salary column পাওয়া যায়নি। Column heading পরীক্ষা করুন।')
                else: st.session_state.df=new; st.success('Gross Salary এবং Net Salary হিসাব করা হয়েছে।'); st.rerun()
        with right:
            if st.button('🗑️ Clear current data',use_container_width=True): st.session_state.df=pd.DataFrame(); st.rerun()
        st.subheader('Excel download'); st.caption('Download-এর সময় salary হিসাব আপডেট হবে এবং headings Excel-এর প্রথম row-তে লেখা হবে।')
        try:
            export_df=salary_calculations(st.session_state.df)
            st.download_button('⬇️ Download Professional Excel',data=make_excel(export_df),file_name=f'AnalyzeData_{date.today().isoformat()}.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',use_container_width=True)
        except Exception as e: st.error(f'Excel তৈরি করা যায়নি: {e}')
        st.divider(); st.subheader('🤖 Gemini AI Data Assistant')
        question=st.text_area('আপনার প্রশ্ন',placeholder='এই ডেটার summary দিন, ভুল বা missing value খুঁজুন...')
        if st.button('Ask Gemini AI'):
            if question.strip():
                with st.spinner('AI বিশ্লেষণ করছে...'):
                    try: st.session_state.ai_result=gemini_answer(question,st.session_state.df)
                    except Exception as e: st.session_state.ai_result=f'Gemini error: {e}'
            else: st.warning('আগে একটি প্রশ্ন লিখুন।')
        if st.session_state.ai_result: st.markdown(st.session_state.ai_result)
elif page=='📁 Templates':
    st.title('📁 Ready-made templates'); st.write('একটি template বেছে নিয়ে নিজের ডেটা দিয়ে পূরণ করুন।')
    templates={'Employee Salary':['Employee ID','Employee Name','Department','Basic Salary','HRA','Other Allowance','Deduction'],'Student Marks':['Roll','Student Name','Subject','Total Marks','Marks Obtained','Grade'],'Sales Report':['Date','Product','Category','Quantity','Unit Price','Total Sales'],'Attendance':['Date','Employee Name','Department','Status','In Time','Out Time'],'Expense Tracker':['Date','Description','Category','Payment Method','Amount']}
    choice=st.selectbox('Template নির্বাচন করুন',list(templates)); st.dataframe(pd.DataFrame(columns=templates[choice]),use_container_width=True)
    if st.button('Use this template'): st.session_state.df=pd.DataFrame(columns=templates[choice]); st.success('Template তৈরি হয়েছে। AI Excel Generator পেজে যান।')
else:
    st.title('ℹ️ About AnalyzeData.io'); st.write('AnalyzeData.io একটি data-to-Excel workbook builder।')
    st.markdown('**যা করতে পারবেন**\n- CSV, XLSX এবং TSV import\n- Data paste ও cell editing\n- Salary calculation\n- Heading, filter ও formatting-সহ Excel download\n- Gemini API key থাকলে AI data analysis')
    st.info('AI ব্যবহার করতে Streamlit Cloud → App settings → Secrets-এ GEMINI_API_KEY যোগ করুন। API key কখনো app.py-তে লিখবেন না।')
