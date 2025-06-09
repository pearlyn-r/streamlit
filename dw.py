import pandas as pd

# Assuming result_df already exists
# If not, load it: result_df = pd.read_excel("your_combined_file.xlsx")

# 1. Categorize the 'limit' column
def categorize_limit(limit):
    try:
        value = float(str(limit).replace('K', '').replace(',', '').strip())
        if value < 50:
            return '<50K'
        elif 50 <= value <= 80:
            return '50K-80K'
        else:
            return '>80K'
    except:
        return 'Unknown'

result_df['limit_category'] = result_df['limit'].apply(categorize_limit)

# 2. Define fixed column order and CE/PE variance thresholds
limit_order = ['<50K', '50K-80K', '>80K']
group_cols = ['date', 'risk_profile', 'limit_category', 'final_product', 'final_rating']
thresholds = list(range(5, 105, 5))

# 3. Pivot generation function
def generate_pivot_per_date(df, var_col, var_prefix):
    for date in df['date'].dropna().unique():
        sub_df = df[df['date'] == date]
        pivot_data = []

        grouped = sub_df.groupby(group_cols)
        for group_keys, group in grouped:
            row = dict(zip(group_cols, group_keys))
            for t in thresholds:
                row[f'{var_prefix}_{t}'] = group[group[var_col] > t].shape[0]
            pivot_data.append(row)

        pivot_df = pd.DataFrame(pivot_data)
        pivot_df['limit_category'] = pd.Categorical(pivot_df['limit_category'], categories=limit_order, ordered=True)
        pivot_df = pivot_df.sort_values(['risk_profile', 'limit_category', 'final_product', 'final_rating'])

        # Reorder columns
        base_cols = group_cols
        var_cols = [f'{var_prefix}_{t}' for t in thresholds]
        pivot_df = pivot_df[base_cols + var_cols]

        # Save to file
        safe_date = pd.to_datetime(date).strftime('%Y-%m-%d')
        filename = f'pivot_{var_prefix}_{safe_date}.xlsx'
        pivot_df.to_excel(filename, index=False)
        print(f"Saved: {filename}")

# 4. Generate pivot tables for CE and PE variance
generate_pivot_per_date(result_df, var_col='ce_variance', var_prefix='ce')
generate_pivot_per_date(result_df, var_col='pe_variance', var_prefix='pe')
