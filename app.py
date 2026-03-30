import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import tempfile

st.title("POA Generator")

uploaded_file = st.file_uploader("Upload Contracts CSV", type=["csv"])

if uploaded_file:
    df = pd.read_csv(uploaded_file)

    # Clean columns
    df.columns = df.columns.str.strip().str.lower().str.replace(" ", "_")

    # Convert dates
    df['end_date'] = pd.to_datetime(df['end_date'])

    # Quarter logic
    def get_quarter(date):
        m = date.month
        return f"Q{((m-1)//3)+1}"

    df['quarter'] = df['end_date'].apply(get_quarter)

    # Savings (5% - 7%)
    df['low_savings'] = (df['contract_value'] * 0.05).round(2)
    df['high_savings'] = (df['contract_value'] * 0.07).round(2)

    # Confidence logic
    def get_confidence(value):
        if value >= 50000:
            return "High"
        elif value >= 10000:
            return "Medium"
        return "Low"

    df['confidence'] = df['contract_value'].apply(get_confidence)

    # Load template
    template_path = "Updated POA Template - Buyers.xlsx"
    wb = load_workbook(template_path)

    # Scenario logic
    def get_scenario(value):
        if value >= 50000:
            return "🟡 FLAT"
        else:
            return "🟠 DOWNGRADE"

    # Fill sheets
    def fill_sheet(sheet_name, data):
        ws = wb[sheet_name]
        row_num = 5

        for _, row in data.iterrows():
            val = row.get('contract_value', 0)

            ws.cell(row=row_num, column=1).value = row.get('tool_name', '')
            ws.cell(row=row_num, column=2).value = row.get('end_date', '')
            ws.cell(row=row_num, column=3).value = round(val / 1000, 2)
            ws.cell(row=row_num, column=4).value = get_scenario(val)
            ws.cell(row=row_num, column=5).value = round((val * 0.05) / 1000, 2)
            ws.cell(row=row_num, column=6).value = round((val * 0.07) / 1000, 2)
            ws.cell(row=row_num, column=7).value = row.get('confidence', '')

            row_num += 1

    # Populate each quarter
    for q in ["Q1", "Q2", "Q3", "Q4"]:
        fill_sheet(q, df[df['quarter'] == q])

    # Save temp file and download
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        wb.save(tmp.name)

        st.success("POA generated successfully")
        st.download_button(
            label="Download POA",
            data=open(tmp.name, "rb"),
            file_name="POA_Output.xlsx"
        )
