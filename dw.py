import pandas as pd

# STEP 1: Map rating to A/B/Other
def map_rating(rating):
    rating_str = str(rating).strip().upper()
    if rating_str.startswith('A'):
        return 'A'
    elif rating_str.startswith('B'):
        return 'B'
    else:
        return 'Other'

# STEP 2: Assume result_df already exists from previous steps
# Add rating category
result_df['rating_category'] = result_df['final_rating'].apply(map_rating)

# STEP 3: Define grouping and thresholds
group_cols = ['date', 'risk_profile', 'limit_category', 'final_product', 'rating_category']

# Thresholds as percentages
thresholds = [t / 100 for t in range(5, 105, 5)]  # 0.05, 0.10, ..., 1.00
threshold_labels = list(range(5, 105, 5))         # For human-readable column names

# STEP 4: Compute counts
final_rows = []
total_rows_counted = 0
all_group_sizes = []

print("\n=== Debugging Summary ===")

for date in result_df['date'].unique():
    sub_df = result_df[result_df['date'] == date]
    print(f"\nProcessing date: {date} | Rows: {len(sub_df)}")
    
    grouped = sub_df.groupby(group_cols, dropna=False)
    print(f"  Groups found: {len(grouped)}")

    for group_keys, group in grouped:
        row = dict(zip(group_cols, group_keys))
        
        group_size = group.shape[0]
        row['group_size'] = group_size
        all_group_sizes.append(group_size)
        total_rows_counted += group_size

        # Debug: Print the first few groups
        if len(final_rows) < 5:
            print(f"    Group: {group_keys}, Size: {group_size}")

        # Count CE and PE thresholds
        for t, label in zip(thresholds, threshold_labels):
            row[f'ce_{label}'] = (group['ce_variance'] > t).sum()
            row[f'pe_{label}'] = (group['pe_variance'] > t).sum()

        final_rows.append(row)

# STEP 5: Create final DataFrame
final_df = pd.DataFrame(final_rows)

# Sort by columns (optional, without enforcing limit order)
final_df = final_df.sort_values(group_cols)

# Define column order
ce_cols = [f'ce_{label}' for label in threshold_labels]
pe_cols = [f'pe_{label}' for label in threshold_labels]
ordered_cols = group_cols + ['group_size'] + ce_cols + pe_cols
final_df = final_df[ordered_cols]

# STEP 6: Save to Excel
output_file = "merged_ce_pe_variance_summary_percent.xlsx"
final_df.to_excel(output_file, index=False)

# STEP 7: Print summary
print("\n=== Summary Report ===")
print(f"Total groups: {len(final_rows)}")
print(f"Total rows counted in all groups: {total_rows_counted}")
print(f"Original DataFrame total rows: {len(result_df)}")
print(f"Output file saved to: {output_file}")
