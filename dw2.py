import pandas as pd
import numpy as np
from openpyxl import load_workbook

def worksheet_to_dataframe(worksheet):
    data = []
    for row in worksheet.iter_rows(values_only=True):
        data.append(row)

    if not data:
        return pd.DataFrame()

    headers = data[0]
    cleaned_headers = []
    for i, header in enumerate(headers):
        cleaned_headers.append(str(header) if header is not None else f"Column_{i}")

    rows_data = data[1:]
    df = pd.DataFrame(rows_data, columns=cleaned_headers)
    df = df.dropna(how='all')
    return df

def process_excel_file(file_path, k_col_index, l_col_index, i_col_index, j_col_index):
    workbook = load_workbook(file_path, read_only=True, data_only=True)
    sheet_names = workbook.sheetnames[3:]  # Skip first 3 sheets

    combined_dfs = []

    for sheet_name in sheet_names:
        worksheet = workbook[sheet_name]
        df = worksheet_to_dataframe(worksheet)

        if df.empty:
            print(f"Skipping empty sheet: {sheet_name}")
            continue

        df['date'] = sheet_name
        col_names = df.columns.tolist()

        # Handle final_rating
        if k_col_index < len(col_names) and l_col_index < len(col_names):
            k_col = col_names[k_col_index]
            l_col = col_names[l_col_index]
            df['final_rating'] = df[k_col].where(
                df[k_col].notna() & (df[k_col].astype(str).str.strip() != ''),
                df[l_col]
            )
        else:
            df['final_rating'] = None
            print(f"Warning: final_rating indices out of range for sheet {sheet_name}")

        # Handle final_product with debugging
        if i_col_index < len(col_names) and j_col_index < len(col_names):
            i_col = col_names[i_col_index]
            j_col = col_names[j_col_index]

            # Debug: Print values for first row
            if not df.empty:
                first_row = df.iloc[0]
                i_val = first_row[i_col]
                j_val = first_row[j_col]
                print(f"\n[DEBUG] Sheet: {sheet_name}")
                print(f"i_col ({i_col_index} - '{i_col}'): '{i_val}'")
                print(f"j_col ({j_col_index} - '{j_col}'): '{j_val}'")
                print(f"  i_val isna: {pd.isna(i_val)}")
                print(f"  i_val is blank/empty: {str(i_val).strip() == ''}")
                final_val = j_val if pd.isna(i_val) or str(i_val).strip() == '' else i_val
                print(f"  Final product value will be: '{final_val}'")

            df['final_product'] = df[i_col].where(
                df[i_col].notna() & (df[i_col].astype(str).str.strip() != ''),
                df[j_col]
            )
        else:
            df['final_product'] = None
            print(f"Warning: final_product indices out of range for sheet {sheet_name}")

        combined_dfs.append(df)

    workbook.close()
    if combined_dfs:
        final_df = pd.concat(combined_dfs, ignore_index=True)
        print(f"\n[INFO] Final combined DataFrame shape: {final_df.shape}")
        return final_df
    else:
        print("No data combined.")
        return pd.DataFrame()
