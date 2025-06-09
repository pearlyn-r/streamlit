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

def is_effectively_null(val):
    """Helper to check if a value should be considered null."""
    if pd.isna(val):
        return True
    val_str = str(val).strip().upper()
    return val_str in ['', '-', 'NULL', '[NULL]']

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

            def get_final_rating(row):
                k_val = row[k_col]
                l_val = row[l_col]
                return l_val if is_effectively_null(k_val) else k_val

            df['final_rating'] = df.apply(get_final_rating, axis=1)
        else:
            df['final_rating'] = None
            print(f"Warning: final_rating indices out of range for sheet {sheet_name}")

        # Handle final_product
        if i_col_index < len(col_names) and j_col_index < len(col_names):
            i_col = col_names[i_col_index]
            j_col = col_names[j_col_index]

            def get_final_product(row):
                i_val = row[i_col]
                j_val = row[j_col]

                # Debug for first row
                if row.name == 0:
                    print(f"\n[DEBUG] Sheet: {sheet_name}")
                    print(f"i_val: '{i_val}', j_val: '{j_val}'")
                    print(f"i_val is_effectively_null: {is_effectively_null(i_val)}")

                return j_val if is_effectively_null(i_val) else i_val

            df['final_product'] = df.apply(get_final_product, axis=1)
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

# === USAGE EXAMPLE ===
if __name__ == "__main__":
    file_path = "your_excel_file.xlsx"  # 🔁 Replace with your file path
    k_col_index = 2  # e.g., col C
    l_col_index = 3  # e.g., col D
    i_col_index = 4  # e.g., col E
    j_col_index = 5  # e.g., col F

    try:
        df = process_excel_file(file_path, k_col_index, l_col_index, i_col_index, j_col_index)
        print("\nFirst 5 rows of result:")
        print(df.head())
    except FileNotFoundError:
        print(f"File not found: {file_path}")
    except Exception as e:
        print(f"Error: {str(e)}")
