import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="GlobXpay Issuance Tool", layout="centered")
st.title("📋 GlobXpay Issuance - Excel to CSC Converter")

# Column index mapping for mandatory fields
MANDATORY_FIELDS = {
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

uploaded_file = st.file_uploader("📤 Upload your Excel sheet", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file, dtype=str)
        df = df.fillna("").astype(str)

        # Ensure all rows have exactly 76 columns
        df = df.iloc[:, :76]
        column_count = df.shape[1]

        if column_count < 76:
            st.error(f"❌ Uploaded file only has {column_count} columns. Expected 76.")
        else:
            validation_errors = []
            for idx, row in df.iterrows():
                missing_fields = []
                for col_idx, field_name in MANDATORY_FIELDS.items():
                    if row[col_idx].strip() == "":
                        missing_fields.append(field_name)
                if missing_fields:
                    validation_errors.append(f"❌ Row {idx + 2} is missing: {', '.join(missing_fields)}")

            if validation_errors:
                st.error("Some rows are missing mandatory fields:")
                for error in validation_errors:
                    st.write(error)
            else:
                st.success(f"✅ All {len(df)} rows passed validation!")

                # Clean and protect leading zeros by wrapping in ="..."
                def protect_text(value):
                    value = value.replace("-", "").strip()
                    return f'="{value}"' if value else ""

                df_cleaned = df.applymap(protect_text)
                csv = df_cleaned.to_csv(index=False, header=False, sep=";").encode("utf-8")
                file_name = f"CSC_Converted_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"

                st.download_button(
                    label="📥 Download CSC CSV File",
                    data=csv,
                    file_name=file_name,
                    mime="text/csv"
                )
    except Exception as e:
        st.error(f"❌ Error processing file: {e}")
