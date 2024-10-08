import pandas as pd
import streamlit as st
from io import BytesIO

# def transform_excel_to_csv(excel_data, sheet_name):
#     df_excel = pd.read_excel(excel_data, sheet_name=sheet_name, keep_default_na=False)
#     df_excel = df_excel.drop([0, 1, 2], errors='ignore')

#     selected_columns = {
#         'Unnamed: 1': 'Midstream_System',
#         'Unnamed: 2': 'ENTITY',
#         'Unnamed: 3': 'January',
#         'Unnamed: 4': 'February',
#         'Unnamed: 5': 'March',
#         'Unnamed: 6': 'April',
#         'Unnamed: 7': 'May',
#         'Unnamed: 8': 'June',
#         'Unnamed: 9': 'July',
#         'Unnamed: 10': 'August',
#         'Unnamed: 11': 'September',
#         'Unnamed: 12': 'October',
#         'Unnamed: 13': 'November',
#         'Unnamed: 14': 'December'
#     }

#     df_transformed = df_excel[list(selected_columns.keys())].rename(columns=selected_columns)
#     df_transformed.insert(0, 'Year', sheet_name)

#     df_transformed = df_transformed.dropna(subset=['Midstream_System', 'ENTITY'])

#     return df_transformed

def transform_excel_to_csv(excel_data, sheet_name):
    df_excel = pd.read_excel(excel_data, sheet_name=sheet_name, keep_default_na=False)
    df_excel = df_excel.drop([0, 1, 2], errors='ignore')

    selected_columns = {
        'Unnamed: 1': 'Midstream_System',
        'Unnamed: 2': 'ENTITY',
        'Unnamed: 3': 'January',
        'Unnamed: 4': 'February',
        'Unnamed: 5': 'March',
        'Unnamed: 6': 'April',
        'Unnamed: 7': 'May',
        'Unnamed: 8': 'June',
        'Unnamed: 9': 'July',
        'Unnamed: 10': 'August',
        'Unnamed: 11': 'September',
        'Unnamed: 12': 'October',
        'Unnamed: 13': 'November',
        'Unnamed: 14': 'December'
    }

    df_transformed = df_excel[list(selected_columns.keys())].rename(columns=selected_columns)
    df_transformed.insert(0, 'Year', sheet_name)

    # Locating the condition
    df_final = df_transformed[(df_transformed['Midstream_System'] .eq("Bracky Branch")) & (df_transformed['ENTITY'] .eq ("CHEMARD"))]
    if not df_final.empty:
        cut_off_index = df_final.index[0]  # Get the first index where the condition is met
        df_transformed = df_transformed.loc[:cut_off_index]  # Keep rows only before the cut-off index

    return df_transformed


def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    processed_data = output.getvalue()
    return processed_data

def main():
    st.title("Excel to CSV Transformer")
    uploaded_file = st.file_uploader("Choose an Excel file", type=['xlsx'])

    if uploaded_file is not None:
        xl = pd.ExcelFile(uploaded_file)
        sheets = xl.sheet_names
        default_sheet = "2024" if "2024" in sheets else sheets[0]
        sheet_name = st.selectbox("Select Sheet:", sheets, index=sheets.index(default_sheet))

        if st.button('Process Sheet'):
            df_transformed = transform_excel_to_csv(uploaded_file, sheet_name)
            st.dataframe(df_transformed.head())  
            
            # Convert DataFrame to Excel for download
            result = to_excel(df_transformed)
            st.download_button(label="Download Excel File",
                               data=result,
                               file_name="transformed_output.csv",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == "__main__":
    main()
