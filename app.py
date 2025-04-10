# Excel Consolidator GPT - Streamlit App Version

import pandas as pd
import os
import re
import streamlit as st
from io import BytesIO

class ExcelConsolidator:
    def __init__(self, files, keywords=None):
        self.files = files
        self.combined_df = pd.DataFrame()
        self.keywords = keywords if keywords else []

    def read_excel_files(self):
        for uploaded_file in self.files:
            try:
                xls = pd.ExcelFile(uploaded_file)
                for sheet in xls.sheet_names:
                    df = xls.parse(sheet)
                    df['source_file'] = uploaded_file.name
                    df['sheet_name'] = sheet
                    df = self.search_keywords(df)
                    self.combined_df = pd.concat([self.combined_df, df], ignore_index=True)
            except Exception as e:
                st.warning(f"Error reading {uploaded_file.name}: {e}")

    def search_keywords(self, df):
        if not self.keywords:
            return df

        keyword_pattern = '|'.join([re.escape(k) for k in self.keywords])
        match_df = df.apply(lambda row: row.astype(str).str.contains(keyword_pattern, case=False).any(), axis=1)
        return df[match_df]

    def clean_and_standardize(self):
        self.combined_df.dropna(how='all', inplace=True)
        self.combined_df.columns = [col.strip().lower().replace(' ', '_') for col in self.combined_df.columns]

    def deduplicate(self):
        self.combined_df.drop_duplicates(inplace=True)

    def consolidate(self):
        self.read_excel_files()
        self.clean_and_standardize()
        self.deduplicate()
        return self.combined_df


def main():
    st.title("📊 Excel Consolidator GPT")
    st.markdown("Upload multiple Excel files and consolidate them using smart keyword search across all sheets.")

    uploaded_files = st.file_uploader("Upload Excel files", type=[".xlsx", ".xls"], accept_multiple_files=True)
    keywords_input = st.text_input("Enter keywords (comma-separated)", "revenue, client name, AUM, performance")
    keywords = [k.strip() for k in keywords_input.split(",") if k.strip()]

    if st.button("🔍 Consolidate") and uploaded_files:
        consolidator = ExcelConsolidator(uploaded_files, keywords)
        result_df = consolidator.consolidate()

        st.success(f"Consolidation complete. {len(result_df)} rows extracted.")
        st.dataframe(result_df.head(50))

        # Download output
        towrite = BytesIO()
        result_df.to_excel(towrite, index=False, engine='openpyxl')
        towrite.seek(0)
        st.download_button("📥 Download Excel", data=towrite, file_name="consolidated_output.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


if __name__ == "__main__":
    main()
