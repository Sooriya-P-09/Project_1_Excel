import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

def merge_files(file1, file2, file3, selected_columns):
    # Read files correctly
    df1 = pd.read_excel(file1, engine="openpyxl") if file1.name.endswith(('.xls', '.xlsx')) else pd.read_csv(file1)
    df2 = pd.read_excel(file2, engine="openpyxl") if file2.name.endswith(('.xls', '.xlsx')) else pd.read_csv(file2)
    df3 = pd.read_excel(file3, engine="openpyxl") if file3.name.endswith(('.xls', '.xlsx')) else pd.read_csv(file3)
    
    # Merge on 'Employee ID'
    merged_df = df1.merge(df3, on='Employee ID', how='left').merge(df2, on='Employee ID', how='left')
    
    # Keep selected columns
    merged_df = merged_df[['Employee ID'] + selected_columns]
    
    # Remove duplicate PDF Reference Numbers
    if 'PDF Reference Number' in merged_df.columns:
        merged_df = merged_df.drop_duplicates(subset=['Employee ID', 'PDF Reference Number'])
    
    # Add certification index
    merged_df['Cert_Index'] = merged_df.groupby('Employee ID').cumcount() + 1
    
    # Pivot table approach for safety
    final_df = merged_df.pivot_table(index='Employee ID', columns='Cert_Index', values=selected_columns, aggfunc='first', fill_value='')
    
    # Flatten multi-index columns
    final_df.columns = [f"{col[0]}_{col[1]}" for col in final_df.columns]
    final_df = final_df.reset_index()
    
    return final_df

# Streamlit App
st.title("File Merger & Pivot Tool")

uploaded_files = st.file_uploader("Upload 3 files", accept_multiple_files=True, type=["csv", "xls", "xlsx"])

if uploaded_files and len(uploaded_files) == 3:
    file1, file2, file3 = uploaded_files

    # Read column names from all files
    df1 = pd.read_excel(file1, engine="openpyxl") if file1.name.endswith(('.xls', '.xlsx')) else pd.read_csv(file1)
    df2 = pd.read_excel(file2, engine="openpyxl") if file2.name.endswith(('.xls', '.xlsx')) else pd.read_csv(file2)
    df3 = pd.read_excel(file3, engine="openpyxl") if file3.name.endswith(('.xls', '.xlsx')) else pd.read_csv(file3)

    all_columns = list(set(df1.columns.tolist() + df2.columns.tolist() + df3.columns.tolist()))

    # Column selection
    selected_columns = st.multiselect("Choose Columns to Process", all_columns, default=[])

    if selected_columns:
        merged_output = merge_files(file1, file2, file3, selected_columns)

        # Display merged output
        st.write("### Merged & Processed Output")
        st.dataframe(merged_output)

        # Convert to Excel in-memory
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            merged_output.to_excel(writer, index=False)
        output.seek(0)

        st.download_button("Download Processed File", output, file_name="merged_output.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
