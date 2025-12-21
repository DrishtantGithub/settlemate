# SettleMate

Streamlit app to summarize UPI settlement sheets:
- Upload multiple Excel files (or a ZIP).
- Sums Fee and Transaction Amount (debit/credit).
- Converts .xls -> .xlsx automatically (if required).
- Produces per-file and date-cycle aggregated CSV exports.

## Run locally
```bash
python -m venv venv
source venv/bin/activate   # Windows: .\venv\Scripts\Activate.ps1
pip install -r requirements.txt
streamlit run streamlit_app.py
