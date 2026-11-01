# AnalyzeData.io — AI Excel Generator (Starter Release)

## Included in this first release
- Blue/white responsive Streamlit dashboard
- Universal prompt box and starter templates
- Employee, Salary, Attendance, and Student Marks structures
- CSV and XLSX import
- Editable spreadsheet preview
- Professional XLSX export (dark-blue header, borders, alternating rows, filters, freeze panes, auto-width, print setup)

## Not included yet
Gemini AI, PDF extraction, image OCR, and automatic AI formula generation are intentionally not shown as working features. They will be integrated and tested in the next development step.

## Run locally
1. Install Python 3.10 or newer.
2. Open a terminal in this folder.
3. Run: `pip install -r requirements.txt`
4. Run: `streamlit run app.py`

## Deploy to Streamlit Community Cloud
1. Upload `app.py` and `requirements.txt` to a GitHub repository.
2. Create a new app at https://share.streamlit.io/
3. Select the repository, branch, and `app.py`.
4. Deploy.

Do not commit API keys or passwords to GitHub. Gemini secrets will be added in a later step using Streamlit Secrets.
