import streamlit as st
import pandas as pd
import openpyxl

def merge_files(file1, file2, file3, selected_columns):
    # Read files (supporting multiple formats)
    df1 = pd.read_excel(file1) if file1.name.endswith('.xls') or file1.name.endswith('.xlsx') else pd.read_csv(file1)
    df2 = pd.read_excel(file2) if file2.name.endswith('.xls') or file2.name.endswith('.xlsx') else pd.read_csv(file2)
    df3 = pd.read_excel(file3) if file3.name.endswith('.xls') or file3.name.endswith('.xlsx') else pd.read_csv(file3)
    
    # Merge files based on common 'Employee ID'
    merged_df = df1.merge(df3, on='Employee ID', how='left').merge(df2, on='Employee ID', how='left')
    
    # Filter only selected columns
    merged_df = merged_df[['Employee ID'] + selected_columns]
    
    # Drop duplicate PDF Reference Numbers for each employee
    if 'PDF Reference Number' in merged_df.columns:
        merged_df = merged_df.drop_duplicates(subset=['Employee ID', 'PDF Reference Number'])
    
    # Generate a unique row number for each certificate per employee
    merged_df['Cert_Index'] = merged_df.groupby('Employee ID').cumcount() + 1
    
    # Pivot to ensure each employee has a single row
    final_df = merged_df.pivot(index='Employee ID', columns='Cert_Index', values=selected_columns)
    
    # Flatten the multi-level column names and rename them properly
    final_df.columns = [f'{col[0]}_{col[1]}' for col in final_df.columns]
    
    # Reset index to make 'Employee ID' a column again
    final_df = final_df.reset_index()
    
    # Reorder columns in the desired sequence
    column_order = ['Employee ID']
    max_cert_index = merged_df['Cert_Index'].max()  # Get the maximum Cert_Index value
    for i in range(1, max_cert_index + 1):
        for col in selected_columns:
            column_order.append(f'{col}_{i}')
    
    # Ensure only existing columns are included
    column_order = [col for col in column_order if col in final_df.columns]
    
    # Reorder the DataFrame columns
    final_df = final_df[column_order]
    
    return final_df

st.title("File Merger & Pivot Tool")

uploaded_files = st.file_uploader("Upload 3 files", accept_multiple_files=True, type=["csv", "xls", "xlsx"])

if uploaded_files and len(uploaded_files) == 3:
    file1, file2, file3 = uploaded_files
    
    # Read all columns from the three files
    df1 = pd.read_excel(file1) if file1.name.endswith('.xls') or file1.name.endswith('.xlsx') else pd.read_csv(file1)
    df2 = pd.read_excel(file2) if file2.name.endswith('.xls') or file2.name.endswith('.xlsx') else pd.read_csv(file2)
    df3 = pd.read_excel(file3) if file3.name.endswith('.xls') or file3.name.endswith('.xlsx') else pd.read_csv(file3)
    
    all_columns = list(set(df1.columns.tolist() + df2.columns.tolist() + df3.columns.tolist()))
    
    # Default empty selection
    selected_columns = st.multiselect("Choose Columns to Process", all_columns, default=[])
    
    if selected_columns:
        merged_output = merge_files(file1, file2, file3, selected_columns)
        
        # Display merged output
        st.write("### Merged & Processed Output")
        st.dataframe(merged_output)
        
        # Download option
        output_file = "merged_output.xlsx"
        merged_output.to_excel(output_file, index=False)
        
        with open(output_file, "rb") as f:
            st.download_button("Download Processed File", f, file_name=output_file, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
