import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="GlobXpay Issuance Tool", layout="centered")

st.title("📋 GlobXpay Issuance - Excel to CSC Converter")

uploaded_file = st.file_uploader("📤 Upload your Excel sheet", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)

        st.success(f"✅ Loaded {len(df)} rows successfully!")

        # Process Data: remove hyphens and keep 76 fields
        df_processed = df.iloc[:, :76].copy()
        df_processed = df_processed.fillna("").astype(str)
        df_processed = df_processed.applymap(lambda x: x.replace("-", "").strip())

        # Preview
        st.write("🔍 Preview (first 5 rows):")
        st.dataframe(df_processed.head())

        # Generate downloadable CSV
        csv = df_processed.to_csv(index=False, header=False, sep=";").encode("utf-8")
        file_name = f"CSC_Converted_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"

        st.download_button(
            label="📥 Download CSC CSV File",
            data=csv,
            file_name=file_name,
            mime="text/csv"
        )

    except Exception as e:
        st.error(f"❌ Error reading file: {e}")
