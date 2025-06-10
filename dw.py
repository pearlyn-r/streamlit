import pandas as pd

# Define thresholds from 5 to 100 with step of 5
thresholds = list(range(5, 105, 5))

def generate_pivot_per_date(df, var_col, var_prefix):
    print(f"\nGenerating pivot for: {var_col}\n{'='*50}")

    df = df.copy()

    # Ensure variance column is numeric
    df[var_col] = pd.to_numeric(df[var_col], errors='coerce')

    # Convert 'date' to string (no datetime)
    df['date'] = df['date'].astype(str)

    # Process each unique date
    for date in df['date'].unique():
        sub_df = df[df['date'] == date]
        print(f"\nDate: {date} — Rows: {len(sub_df)}")

        pivot_data = []
        grouped = sub_df.groupby(['date', 'risk_profile', 'limit', 'final_product', 'final_rating'], dropna=False)
        print(f"  Groups found: {len(grouped)}")

        for group_keys, group_df in grouped:
            row = dict(zip(['date', 'risk_profile', 'limit', 'final_product', 'final_rating'], group_keys))
            row['group_size'] = len(group_df)

            for t in thresholds:
                count = group_df[group_df[var_col] > t].shape[0]
                row[f'{var_prefix}_{t}'] = count

            pivot_data.append(row)

        pivot_df = pd.DataFrame(pivot_data)

        # Order columns
        base_cols = ['date', 'risk_profile', 'limit', 'final_product', 'final_rating', 'group_size']
        var_cols = [f'{var_prefix}_{t}' for t in thresholds]
        pivot_df = pivot_df[base_cols + var_cols]

        # Save
        safe_date = date.replace('/', '-').replace('\\', '-').replace(':', '-')
        filename = f'pivot_{var_prefix}_{safe_date}.xlsx'
        pivot_df.to_excel(filename, index=False)
        print(f"Saved: {filename}")
