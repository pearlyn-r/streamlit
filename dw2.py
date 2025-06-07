import pandas as pd
import numpy as np
from openpyxl import load_workbook

def process_excel_file(file_path, k_col_index, l_col_index, i_col_index, j_col_index):
    """
    Process Excel file by combining sheets with additional columns.
    
    Parameters:
    - file_path: path to Excel file
    - k_col_index: column index for final_rating (0-based, e.g., 0 for column A)
    - l_col_index: column index for final_rating (0-based, e.g., 1 for column B)
    - i_col_index: column index for final_product (0-based)
    - j_col_index: column index for final_product (0-based)
    """
    
    # Load workbook using openpyxl
    workbook = load_workbook(file_path, read_only=True)
    sheet_names = workbook.sheetnames
    
    print(f"Found {len(sheet_names)} sheets: {sheet_names}")
    
    # Skip first 3 sheets and process the rest
    sheets_to_process = sheet_names[3:]
    print(f"Processing sheets: {sheets_to_process}")
    
    combined_dfs = []
    
    for sheet_name in sheets_to_process:
        # Load the sheet as DataFrame using openpyxl engine
        df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')
        
        # Add date column with sheet name
        df['date'] = sheet_name
        
        # Get column names for the specified indices
        col_names = df.columns.tolist()
        
        # Create final_rating column
        if k_col_index < len(col_names) and l_col_index < len(col_names):
            k_col_name = col_names[k_col_index]
            l_col_name = col_names[l_col_index]
            
            # Take value from k_col if not null, otherwise from l_col
            df['final_rating'] = df[k_col_name].fillna(df[l_col_name])
        else:
            print(f"Warning: Column indices {k_col_index} or {l_col_index} out of range for sheet {sheet_name}")
            df['final_rating'] = None
        
        # Create final_product column
        if i_col_index < len(col_names) and j_col_index < len(col_names):
            i_col_name = col_names[i_col_index]
            j_col_name = col_names[j_col_index]
            
            # Take value from i_col if not null, otherwise from j_col
            df['final_product'] = df[i_col_name].fillna(df[j_col_name])
        else:
            print(f"Warning: Column indices {i_col_index} or {j_col_index} out of range for sheet {sheet_name}")
            df['final_product'] = None
        
        combined_dfs.append(df)
        print(f"Processed sheet '{sheet_name}' with {len(df)} rows")
    
    # Close the workbook
    workbook.close()
    
    # Combine all DataFrames
    if combined_dfs:
        final_df = pd.concat(combined_dfs, ignore_index=True)
        print(f"Combined DataFrame created with {len(final_df)} total rows")
        return final_df
    else:
        print("No sheets to process")
        return pd.DataFrame()

# Example usage
if __name__ == "__main__":
    # CUSTOMIZE THESE PARAMETERS:
    file_path = "your_excel_file.xlsx"  # Replace with your Excel file path
    
    # Column indices (0-based). For example:
    # Column A = 0, Column B = 1, Column C = 2, etc.
    k_col_index = 2  # Replace with actual column index for final_rating
    l_col_index = 3  # Replace with actual column index for final_rating
    i_col_index = 4  # Replace with actual column index for final_product
    j_col_index = 5  # Replace with actual column index for final_product
    
    try:
        # Process the Excel file
        result_df = process_excel_file(file_path, k_col_index, l_col_index, i_col_index, j_col_index)
        
        # Display basic info about the result
        print("\nFinal DataFrame Info:")
        print(f"Shape: {result_df.shape}")
        print(f"Columns: {list(result_df.columns)}")
        
        # Display first few rows
        print("\nFirst 5 rows:")
        print(result_df.head())
        
        # Save to new Excel file
        output_file = "combined_data.xlsx"
        result_df.to_excel(output_file, index=False)
        print(f"\nCombined data saved to: {output_file}")
        
    except FileNotFoundError:
        print(f"Error: File '{file_path}' not found")
    except Exception as e:
        print(f"Error processing file: {str(e)}")

# Alternative version with column names instead of indices
def process_excel_file_by_names(file_path, k_col_name, l_col_name, i_col_name, j_col_name):
    """
    Alternative version that uses column names instead of indices.
    """
    workbook = load_workbook(file_path, read_only=True)
    sheet_names = workbook.sheetnames[3:]  # Skip first 3 sheets
    
    combined_dfs = []
    
    for sheet_name in sheet_names:
        df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')
        df['date'] = sheet_name
        
        # Create final_rating column using column names
        if k_col_name in df.columns and l_col_name in df.columns:
            df['final_rating'] = df[k_col_name].fillna(df[l_col_name])
        else:
            df['final_rating'] = None
            
        # Create final_product column using column names
        if i_col_name in df.columns and j_col_name in df.columns:
            df['final_product'] = df[i_col_name].fillna(df[j_col_name])
        else:
            df['final_product'] = None
            
        combined_dfs.append(df)
    
    workbook.close()
    return pd.concat(combined_dfs, ignore_index=True) if combined_dfs else pd.DataFrame()
