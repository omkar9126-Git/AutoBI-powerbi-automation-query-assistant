
import streamlit as st
import pandas as pd
from query_parser import process_query
from utils import get_autocomplete_options, render_chart

st.title("📊 Excel Query Chatbot with Natural Language")

st.markdown("Upload one or more Excel files and ask questions like:")
st.code("Heat - N6491, Material = Steel, Quantities > 10", language="text")

uploaded_files = st.file_uploader("Upload Excel file(s)", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    all_data = []
    sheet_info = {}

    for file in uploaded_files:
        xls = pd.ExcelFile(file)
        for sheet in xls.sheet_names:
            df = xls.parse(sheet)
            df['__source_file__'] = file.name
            df['__sheet__'] = sheet
            all_data.append(df)
            sheet_info[f"{file.name} - {sheet}"] = df

    combined_df = pd.concat(all_data, ignore_index=True)
    st.session_state["data"] = combined_df

    st.success("Files uploaded and combined!")

    st.subheader("💬 Ask your query:")
    autocomplete_options = get_autocomplete_options(combined_df)

    col1, col2, col3 = st.columns(3)
    heat_input = col1.selectbox("Heat", options=[""] + autocomplete_options["heats"])
    material_input = col2.selectbox("Material", options=[""] + autocomplete_options["materials"])
    section_input = col3.selectbox("Section", options=[""] + autocomplete_options["sections"])

    query_input = st.text_input("Type additional conditions (e.g., Quantities > 10, Summary)", key="query")

    full_query = []
    if heat_input:
        full_query.append(f"Heat - {heat_input}")
    if material_input:
        full_query.append(f"Material = {material_input}")
    if section_input:
        full_query.append(f"Section = {section_input}")
    if query_input:
        full_query.append(query_input)

    combined_query = ", ".join(full_query)

    if combined_query:
        filtered_df, summary_or_chart = process_query(combined_query, combined_df)
        if not filtered_df.empty:
            st.dataframe(filtered_df, use_container_width=True)
        if summary_or_chart:
            render_chart(summary_or_chart)
        st.download_button("Download Filtered Data as CSV", data=filtered_df.to_csv(index=False), file_name="filtered_data.csv")

else:
    st.warning("Please upload at least one Excel file.")
