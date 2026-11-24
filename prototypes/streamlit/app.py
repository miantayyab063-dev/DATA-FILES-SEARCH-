import os
import pandas as pd
import streamlit as st

st.set_page_config(page_title="File Data Searcher", layout="wide")
st.title("Excel & CSV Deep Search Engine")

# Files wala folder path
FOLDER_PATH = "data_files"  # Apne folder ka naam yahan likhein

search_query = st.text_input("Enter search word / keyword:")

if search_query and os.path.exists(FOLDER_PATH):
    found_any = False
    files = [f for f in os.listdir(FOLDER_PATH) if f.endswith(('.xlsx', '.xls', '.csv'))]

    for file_name in files:
        file_path = os.path.join(FOLDER_PATH, file_name)
        
        try:
            # CSV file handling
            if file_name.endswith('.csv'):
                sheets_data = {"Sheet1": pd.read_csv(data)}
            # Excel file handling (sari sheets read karna)
            else:
                sheets_data = pd.read_excel(file_path, sheet_name=None)
            
            for sheet_name, df in sheets_data.items():
                if df.empty:
                    continue
                
                # Pure dataframe ko text mein convert kar ke search karna
                mask = df.astype(str).apply(
                    lambda row: row.str.contains(search_query.strip(), case=False, na=False).any(),
                    axis=1
                )
                
                matched_rows = df[mask]
                
                # Agar match mil jaye toh display karein
                if not matched_rows.empty:
                    found_any = True
                    st.success(f"Match found in: **{file_name}** | Sheet: **{sheet_name}** ({len(matched_rows)} rows)")
                    st.dataframe(matched_rows, use_container_width=True)
                    st.divider()

        except Exception as e:
            st.error(f"Error reading {file_name}: {str(e)}")

    if not found_any:
        st.warning("Diye gaye lafz ka koi record kisi file mein nahi mila.")