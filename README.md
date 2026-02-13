# Budget File Report Converter (Streamlit)
## What it does
- Drops first 9 rows and the original row 11
- Uses the first remaining row as header
- Extracts Budget_Code / Budget_Type from group rows:
  - XXXX : something
  - XXXX : something
- Forward-fills group info to item rows
- Splits item line into Expense_Code + Expense_Detail
- Provides a downloadable converted .xlsx
