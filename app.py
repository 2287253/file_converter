# imports
import streamlit as st
import pandas as pd
import os
import base64
from io import BytesIO
from docx import Document
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle

# Set up our App
st.set_page_config(page_title="Data sweeper", layout='wide')
st.title("Data sweeper")
st.write("Transform your files between CSV and Excel formats with built-in data cleaning and visualization!")

# Upload file
uploaded_files = st.file_uploader("Upload your files (CSV, Excel, Word, PDF):", type=["csv", "xlsx", "xls", "doc", "docx", "pdf"],
accept_multiple_files=True)

# Check if files were uploaded
if uploaded_files:
    for file in uploaded_files:
        file_ext = os.path.splitext(file.name)[-1].lower()

        if file_ext == ".csv":
            df = pd.read_csv(file)
        elif file_ext in [".xlsx", ".xls"]:
            df = pd.read_excel(file)
        elif file_ext in [".doc", ".docx"]:
            st.error("Word files can only be converted to PDF. Please use a Word file with tables.")
            continue
        elif file_ext == ".pdf":
            st.error("PDF files can only be converted to other formats. Please use a PDF with tables.")
            continue
        else:
            st.error(f"Unsupported file format: {file_ext}")
            continue

        # Display the uploaded file
        st.write(f"**File Name:** {file.name}")
        file_size = file.size
        st.write(f"**File Size:** {file_size/1290}mb")

        # Shows 5 Rows of out data file
        st.write("preview the Head of the dataframe")
        st.dataframe(df.head())
        
        #Options for Data Cleaning 
        st.subheader("Data Cleaning Options")
        if st.checkbox(f"Clean data for {file.name}"):
            col1, col2 = st.columns(2)

            with col1:
                if st.button(f"Remove Duplicates from {file.name}"):
                    df.drop_duplicates(inplace=True)
                    st.write("Duplicates Removed!")

            with col2:
                if st.button(f"Fill Missing Values for {file.name}"):
                    numeric_cols = df.select_dtypes(include=['number']).columns
                    df[numeric_cols] = df[numeric_cols].fillna(df[numeric_cols].mean())
                    st.write("Missing Values have been Filled!")

            # Choose Specific Columns to keep or Convert
            st.subheader("Select Columns to Convert")
            columns = st.multiselect(f"Choose Columns for {file.name}", df.columns, default=df.columns)
            df = df[columns]

            # Create Some Visualizations 
            st.subheader("Data Visualization")
            if st.checkbox(f"Show Visualization for {file.name}"):
                st.bar_chart(df.select_dtypes(include='number').iloc[:,:2])

                # Convert the file -> csv to Excel
                st.subheader("Conversion Options")
                conversion_type = st.radio(f"Convert {file.name} to:", ["CSV", "Excel", "PDF", "Word"], key=f"conv_{file.name}")
                if st.button(f"Convert {file.name}"):
                    buffer = BytesIO()
                    if conversion_type == "CSV":
                        df.to_csv(buffer, index=False)
                        file_name = file.name.replace(file_ext, ".csv")
                        mime_type = "text/csv"
                        buffer.seek(0)

                    elif conversion_type == "Excel":
                        df.to_excel(buffer, index=False)
                        file_name = file.name.replace(file_ext, ".xlsx")
                        mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        buffer.seek(0)

                    elif conversion_type == "PDF":
                        # Create PDF using reportlab
                        doc = SimpleDocTemplate(buffer, pagesize=letter)
                        elements = []
                        
                        # Convert DataFrame to list for the table
                        data = [df.columns.tolist()] + df.values.tolist()
                        
                        # Create the table
                        table = Table(data)
                        table.setStyle(TableStyle([
                            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                            ('FONTSIZE', (0, 0), (-1, 0), 12),
                            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                            ('BACKGROUND', (0, 1), (-1, -1), colors.white),
                            ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
                            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                            ('FONTSIZE', (0, 1), (-1, -1), 10),
                            ('GRID', (0, 0), (-1, -1), 1, colors.black)
                        ]))
                        elements.append(table)
                        
                        # Build PDF
                        doc.build(elements)
                        buffer.seek(0)
                        file_name = file.name.replace(file_ext, ".pdf")
                        mime_type = "application/pdf"

                    elif conversion_type == "Word":
                        # Create Word document
                        doc = Document()
                        # Add table to document
                        table = doc.add_table(rows=len(df)+1, cols=len(df.columns))
                        
                        # Add headers
                        for j, column in enumerate(df.columns):
                            table.cell(0, j).text = str(column)
                        
                        # Add data
                        for i, row in enumerate(df.values):
                            for j, value in enumerate(row):
                                table.cell(i+1, j).text = str(value)
                        
                        # Save to buffer
                        doc.save(buffer)
                        buffer.seek(0)
                        file_name = file.name.replace(file_ext, ".docx")
                        mime_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

                    # Download Button
                    st.download_button(
                        label=f"Download {file.name} as {conversion_type}",
                        data=buffer,
                        file_name=file_name,
                        mime=mime_type,
                        key=f"download_{file.name}"
                    )
                    st.success("File processed successfully!")

            # Show the first five rows of the dataframe