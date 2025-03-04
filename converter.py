import streamlit as st
import pandas as pd
import os
from io import BytesIO

st.set_page_config(page_title="Data Sweeper", layout='wide')
st.title("Data Sweeper")
st.write("Transform your files between CSV and Excel formats with built-in data cleaning and visualization!")

# Function to Convert DataFrame to Excel
def convert_to_excel(df, file_name):
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Sheet1")
    buffer.seek(0)
    
    return buffer, f"{file_name}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

# Function to Convert DataFrame to CSV
def convert_to_csv(df, file_name):
    buffer = BytesIO()
    df.to_csv(buffer, index=False)
    buffer.seek(0)

    return buffer, f"{file_name}.csv", "text/csv"

# Upload Files
uploaded_files = st.file_uploader("Upload your files (CSV, Excel):", type=["csv", "xlsx"], accept_multiple_files=True)

if uploaded_files:
    for file in uploaded_files:
        file_ext = os.path.splitext(file.name)[-1].lower()

        # Read Files
        if file_ext == ".csv":
            df = pd.read_csv(file)
        elif file_ext == ".xlsx":
            df = pd.read_excel(file, engine="openpyxl")
        else:
            st.error(f"Unsupported file type: {file_ext}")
            continue

        st.write(f'**File Name:** {file.name}')
        st.write(f"**File Size:** {file.size / 1024:.2f} KB")

        # Preview Data
        st.subheader("Preview of the Data")
        st.dataframe(df.head())

        # Data Cleaning Options
        st.subheader("Data Cleaning Options")
        if st.checkbox(f"Clean Data for {file.name}"):
            col1, col2 = st.columns(2)

            with col1:
                if st.button(f"Remove Duplicates from {file.name}"):
                    df.drop_duplicates(inplace=True)
                    st.write("Duplicates Removed!")

            with col2:
                if st.button(f"Fill Missing Values for {file.name}"):
                    numeric_cols = df.select_dtypes(include=["number"]).columns
                    df[numeric_cols] = df[numeric_cols].fillna(df[numeric_cols].mean())
                    st.write("Missing values have been filled!")

        # Column Selection
        st.subheader("Select Columns to Keep")
        columns = st.multiselect(f"Choose Columns for {file.name}", df.columns, default=df.columns)
        df = df[columns]

        # Data Visualization
        st.subheader("Data Visualization")
        if st.checkbox(f"Show Visualization for {file.name}"):
            st.bar_chart(df.select_dtypes(include='number').iloc[:, :2])

        # Conversion Options
        st.subheader("Conversion Options")
        conversion_type = st.radio(f"Convert {file.name} to:", ["CSV", "Excel"], key=file.name)

        if st.button(f"Convert {file.name}"):
            buffer = BytesIO()
            new_file_name = file.name.rsplit(".", 1)[0]

            if conversion_type == "CSV":
                buffer, file_name, mime_type = convert_to_csv(df, new_file_name)

            elif conversion_type == "Excel":
                buffer, file_name, mime_type = convert_to_excel(df, new_file_name)

            # Download Button
            st.download_button(
                label=f"Download {file_name}",
                data=buffer,
                file_name=file_name,
                mime=mime_type
            )

    st.success("All files processed successfully! 🎉")
