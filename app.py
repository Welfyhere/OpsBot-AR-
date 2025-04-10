# Excel Consolidator GPT - Interactive Dashboard + Visuals + Landing Polish

import pandas as pd
import os
import re
import streamlit as st
from io import BytesIO
import plotly.express as px

class ExcelConsolidator:
    def __init__(self, files, keywords=None):
        self.files = files
        self.combined_df = pd.DataFrame()
        self.keywords = [k.lower().strip().replace(" ", "_") for k in keywords] if keywords else []

    def read_excel_files(self):
        for uploaded_file in self.files:
            try:
                xls = pd.ExcelFile(uploaded_file)
                for sheet in xls.sheet_names:
                    df = xls.parse(sheet)
                    df['source_file'] = uploaded_file.name
                    df['sheet_name'] = sheet
                    filtered_df = self.smart_search(df)
                    if not filtered_df.empty:
                        self.combined_df = pd.concat([self.combined_df, filtered_df], ignore_index=True)
                    else:
                        self.combined_df = pd.concat([self.combined_df, df], ignore_index=True)
            except Exception as e:
                st.warning(f"⚠️ Error reading {uploaded_file.name}: {e}")

    def smart_search(self, df):
        if not self.keywords:
            return df

        df.columns = [c.lower().strip().replace(" ", "_") for c in df.columns]
        matched_cols = [col for col in df.columns if any(term in col for term in self.keywords)]

        text_match = df.apply(lambda row: row.astype(str).str.contains('|'.join(self.keywords), case=False).any(), axis=1)
        matched_rows = df[text_match]

        if matched_cols:
            cols_to_return = [col for col in matched_cols if col in df.columns]
            if 'source_file' in df.columns: cols_to_return.append('source_file')
            if 'sheet_name' in df.columns: cols_to_return.append('sheet_name')
            if matched_rows.empty:
                return df[cols_to_return] if cols_to_return else df
            else:
                return matched_rows[cols_to_return] if cols_to_return else matched_rows

        return matched_rows if not matched_rows.empty else df

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

    def generate_insights(self):
        insights = []
        detailed_summary = []

        if 'revenue' in self.combined_df.columns:
            total_revenue = self.combined_df['revenue'].sum()
            avg_revenue = self.combined_df['revenue'].mean()
            top_revenue = self.combined_df.sort_values(by='revenue', ascending=False).head(1)
            insights.append(f"💰 **Total Revenue**: `${total_revenue:,.2f}`")
            insights.append(f"📈 **Avg Revenue per Entry**: `${avg_revenue:,.2f}`")
            if 'client_name' in top_revenue.columns:
                top_client = top_revenue.iloc[0]['client_name']
                insights.append(f"🏆 **Top Client by Revenue**: `{top_client}`")
            detailed_summary.append(f"Revenue totals `${total_revenue:,.2f}`. Average per row: `${avg_revenue:,.2f}`. Top performer: `{top_client}`.")

        if 'aum' in self.combined_df.columns:
            total_aum = self.combined_df['aum'].sum()
            top_aum = self.combined_df.sort_values(by='aum', ascending=False).head(1)
            insights.append(f"🏦 **Total AUM**: `${total_aum:,.2f}`")
            if 'client_name' in top_aum.columns:
                top_holder = top_aum.iloc[0]['client_name']
                insights.append(f"👑 **Top AUM Holder**: `{top_holder}`")
            detailed_summary.append(f"AUM stands at `${total_aum:,.2f}`. Highest holding from `{top_holder}`.")

        if 'performance' in self.combined_df.columns:
            perf_summary = self.combined_df['performance'].value_counts().to_dict()
            insights.append("📊 **Performance Breakdown**:<br>" + "<br>".join([f"- {k}: {v}" for k,v in perf_summary.items()]))
            top_perf = max(perf_summary, key=perf_summary.get)
            detailed_summary.append(f"Most frequent performance rating is `{top_perf}`. Overall breakdown available above.")

        if 'jurisdiction' in self.combined_df.columns:
            top_juris = self.combined_df['jurisdiction'].value_counts().idxmax()
            count_juris = self.combined_df['jurisdiction'].value_counts().max()
            insights.append(f"🌍 **Most Common Jurisdiction**: `{top_juris}` ({count_juris} entries)")
            detailed_summary.append(f"The jurisdiction `{top_juris}` appears most frequently, with `{count_juris}` entries.")

        if 'call_(x)' in self.combined_df.columns:
            total_calls = self.combined_df['call_(x)'].sum()
            insights.append(f"📞 **Total Calls Logged**: `{total_calls}`")
            detailed_summary.append(f"Across all data, `{total_calls}` calls have been logged.")

        return insights, detailed_summary


def main():
    st.set_page_config(layout="wide", page_title="Excel Consolidator GPT", page_icon="📊")
    st.markdown("""
        <style>
            .block-container { padding-top: 2rem; }
            .stApp { background-color: #0d1117; color: white; font-family: 'Segoe UI', sans-serif; }
            h1, h2, h3 { color: #58a6ff; }
        </style>
    """, unsafe_allow_html=True)

    st.title("📊 Excel Consolidator GPT")
    st.markdown("Effortlessly merge and extract insights from Excel files. Interactive visuals. Business-ready breakdowns.")

    uploaded_files = st.file_uploader("📂 Upload Excel Files", type=[".xlsx", ".xls"], accept_multiple_files=True)
    keywords_input = st.text_input("🔎 Keywords (comma-separated)", "revenue, client, aum, performance")
    keywords = [k.strip() for k in keywords_input.split(",") if k.strip()]

    if st.button("🚀 Run Consolidation") and uploaded_files:
        consolidator = ExcelConsolidator(uploaded_files, keywords)
        result_df = consolidator.consolidate()

        if result_df.empty:
            st.error("No data found with those keywords — showing fallback.")
        else:
            st.success(f"Consolidation complete. Showing {len(result_df)} rows.")

        st.subheader("📋 Full Data Preview")
        st.dataframe(result_df, use_container_width=True)

        insights, summaries = consolidator.generate_insights()

        st.subheader("📌 Key Metrics")
        for ins in insights:
            st.markdown(ins, unsafe_allow_html=True)

        st.subheader("🧠 Executive Summary")
        for s in summaries:
            st.markdown(f"- {s}")

        # 🔢 Visualizations
        if 'revenue' in result_df.columns and 'client_name' in result_df.columns:
            fig = px.bar(result_df.sort_values(by='revenue', ascending=False).head(10), x='client_name', y='revenue', title='Top 10 Clients by Revenue')
            st.plotly_chart(fig, use_container_width=True)

        if 'performance' in result_df.columns:
            perf_chart = px.pie(result_df, names='performance', title='Performance Distribution')
            st.plotly_chart(perf_chart, use_container_width=True)

        if 'jurisdiction' in result_df.columns:
            juris_chart = px.histogram(result_df, x='jurisdiction', title='Jurisdiction Frequency')
            st.plotly_chart(juris_chart, use_container_width=True)

        # 💾 Download
        towrite = BytesIO()
        result_df.to_excel(towrite, index=False, engine='openpyxl')
        towrite.seek(0)
        st.download_button("📥 Download Consolidated Excel", data=towrite, file_name="consolidated_output.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


if __name__ == "__main__":
    main()
