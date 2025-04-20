import streamlit as st
import pandas as pd
from io import StringIO
from datetime import datetime
import csv

st.set_page_config(page_title="GlobXpay Issuance Tool", layout="centered")

st.title("📋 GlobXpay Issuance - Excel to CSC Converter")

uploaded_file = st.file_uploader("📤 Upload your Excel sheet", type=["xlsx"])

# === Mandatory Fields by Column Index ===
mandatory_fields = {
    0: "Record Date",
    1: "Institution Number",
    2: "Branch",
    3: "Application Number",
    4: "Action",
    7: "Customer ID",
    9: "Entity Type",
    11: "First Name",
    13: "Last Name",
    14: "Birth Date",
    16: "Identity Type",
    17: "Identity Number",
    19: "Gender",
    20: "Nationality",
    21: "Phone Number",
    24: "City Code",
    25: "Country Code",
    26: "Address Line 1",
    32: "Bank Account Number",
    34: "Account Name",
    35: "Credit Limit",
    38: "Cardholder Name",
    44: "Product Code",
    53: "ID Expiry Date"
}

# === CLEANING + VALIDATION ===
def validate_and_clean(df):
    errors = []
    cleaned_rows = []

    for idx, row in df.iterrows():
        row_errors = []
        row_values = []

        for i in range(76):
            val = str(row[i]).strip() if i < len(row) and pd.notna(row[i]) else ""

            if i in mandatory_fields and not val:
                row_errors.append(f"Missing '{mandatory_fields[i]}'")

            row_values.append(val)

        if row_errors:
            errors.append(f"❌ Row {idx + 2}: " + ", ".join(row_errors))  # +2 for Excel row number
        cleaned_rows.append(row_values)

    return cleaned_rows, errors

# === CONVERT TO CLEAN CSV ===
def convert_to_csv(data):
    string_buffer = StringIO()
    writer = csv.writer(
        string_buffer,
        delimiter=";",
        quoting=csv.QUOTE_NONE,
        escapechar="\\",
        lineterminator="\n"
    )
    writer.writerows(data)
    return string_buffer.getvalue().encode("utf-8")

# === APP LOGIC ===
if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file, dtype=str).fillna("").astype(str)
        df = df.iloc[:, :76]  # Trim to 76 columns max

        st.success(f"✅ Loaded {len(df)} rows successfully!")

        cleaned_data, errors = validate_and_clean(df)

        if errors:
            st.error("⚠️ Validation Issues Found:")
            for e in errors:
                st.text(e)
        else:
            st.success("✅ All data is valid. Ready to download.")
            csv_data = convert_to_csv(cleaned_data)
            file_name = f"CSC_Converted_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"

            st.download_button(
                label="📥 Download CSC CSV File",
                data=csv_data,
                file_name=file_name,
                mime="text/csv"
            )

    except Exception as e:
        st.error(f"❌ Error processing file:\n{e}")
