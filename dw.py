import pandas as pd

def generate_variance_pivot(df, var_col, label):
    group_cols = ['date', 'risk_profile', 'limit', 'final_product', 'final_rating']
    thresholds = list(range(5, 105, 5))
    result_rows = []

    # Group by unique row identifiers
    for name, group in df.groupby(group_cols):
        row_data = dict(zip(group_cols, name))
        for threshold in thresholds:
            count = group[group[var_col] > threshold].shape[0]
            row_data[f'{label}_{threshold}'] = count
        result_rows.append(row_data)

    return pd.DataFrame(result_rows)

# Generate CE variance pivot table
pivot_ce = generate_variance_pivot(result_df, var_col='ce_variance', label='ce')

# Generate PE variance pivot table
pivot_pe = generate_variance_pivot(result_df, var_col='pe_variance', label='pe')

# Merge both pivot tables on row identifiers
merged_pivot = pd.merge(
    pivot_ce,
    pivot_pe,
    on=['date', 'risk_profile', 'limit', 'final_product', 'final_rating'],
    how='outer'
)

# Save to Excel
merged_pivot.to_excel("variance_threshold_summary_wide.xlsx", index=False)
print("✅ Saved final wide-format summary to 'variance_threshold_summary_wide.xlsx'")
