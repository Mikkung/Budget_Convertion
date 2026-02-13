# Budget File Converter (Streamlit)

## Run locally
1) Create a virtual environment (recommended)
2) Install dependencies:
   pip install -r requirements.txt
3) Start the app:
   streamlit run app.py

## What it does
- Drops first 9 rows and the original row 11
- Uses the first remaining row as header
- Extracts Budget_Code / Budget_Type from group rows:
  - G501 : something
  - G501_1 : something
- Forward-fills group info to item rows
- Splits item line into Expense_Code + Expense_Detail
- Provides a downloadable converted .xlsx
