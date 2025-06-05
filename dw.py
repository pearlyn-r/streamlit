# compare_sql_dataframes.py
import pandas as pd
import pyodbc

def fetch_data(query: str, connection_str: str) -> pd.DataFrame:
    conn = pyodbc.connect(connection_str)
    return pd.read_sql(query, conn)

def compare_dataframes(df1: pd.DataFrame, df2: pd.DataFrame, date1: str, date2: str, exclude_columns: list[str]) -> pd.DataFrame:
    df1.columns = [f"{col}_{date1}" for col in df1.columns]
    df2.columns = [f"{col}_{date2}" for col in df2.columns]

    df_combined = pd.concat([df1, df2], axis=1)

    for col in df1.columns:
        orig_col = col.replace(f"_{date1}", '')
        if orig_col in exclude_columns:
            continue

        col1 = f"{orig_col}_{date1}"
        col2 = f"{orig_col}_{date2}"

        if col1 in df_combined.columns and col2 in df_combined.columns:
            if df_combined[col1].dtype == object or df_combined[col2].dtype == object:
                df_combined[f"{orig_col}_change"] = df_combined.apply(
                    lambda row: f"{row[col1]} -> {row[col2]}" if row[col1] != row[col2] else row[col1], axis=1
                )
            else:
                df_combined[f"{orig_col}_change"] = df_combined[col2] - df_combined[col1]

    return df_combined

# Example usage:
if __name__ == "__main__":
    connection_str = "DRIVER={SQL Server};SERVER=server;DATABASE=db;UID=user;PWD=password"
    query_template = "SELECT * FROM your_table WHERE cob_date = '{}'"

    date1 = "2024-01-01"
    date2 = "2024-02-01"

    df1 = fetch_data(query_template.format(date1), connection_str)
    df2 = fetch_data(query_template.format(date2), connection_str)

    result = compare_dataframes(df1, df2, date1, date2, exclude_columns=["id", "name", "cob_date"])

    result.to_csv("comparison_output.csv", index=False)
    print("Comparison saved to comparison_output.csv")
