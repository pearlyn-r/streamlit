import pandas as pd

# Step 1: Map final_rating to simplified category
def map_rating(rating):
    rating_str = str(rating).strip().upper()
    if rating_str.startswith('A'):
        return 'A'
    elif rating_str.startswith('B'):
        return 'B'
    else:
        return 'Other'

# Step 2: Apply rating category
result_df['rating_category'] = result_df['final_rating'].apply(map_rating)

# Step 3: Define grouping columns and thresholds
group_cols = ['date', 'risk_profile', 'limit', 'final_product', 'rating_category']
thresholds = list(range(5, 105, 5))  # 5 to 100 inclusive

# Step 4: Initialize result storage
final_rows = []
total_group_size = 0
group_count = 0

print("Starting group processing...\n")

# Step 5: Iterate over each date group
for date in result_df['date'].unique():
    sub_df = result_df[result_df['date'] == date]
    print(f"Processing date: {date}, Rows in date: {len(sub_df)}")

    grouped = sub_df.groupby(group_cols, dropna=False)

    for group_keys, group in grouped:
        group_size = len(group)
        total_group_size += group_size
        group_count += 1

        row = dict(zip(group_cols, group_keys))
        row['group_size'] = group_size  # Add group size to row

        print(f"  Group {group_count}: {group_keys} | Size: {group_size}")

        # Check and debug if any variance is missing
        if group['ce_variance'].isnull().any():
            print(f"    WARNING: CE variance has missing values in group {group_keys}")
        if group['pe_variance'].isnull().any():
            print(f"    WARNING: PE variance has missing values in group {group_keys}")

        # Count CE variance threshold exceedances
        for t in thresholds:
            row[f'ce_{t}'] = (group['ce_variance'] > t).sum()

        # Count PE variance threshold exceedances
        for t in thresholds:
            row[f'pe_{t}'] = (group['pe_variance'] > t).sum()

        final_rows.append(row)

# Step 6: Create final DataFrame
final_df = pd.DataFrame(final_rows)

# Step 7: Reorder columns
ce_cols = [f'ce_{t}' for t in thresholds]
pe_cols = [f'pe_{t}' for t in thresholds]
ordered_cols = group_cols + ['group_size'] + ce_cols + pe_cols
final_df = final_df[ordered_cols]

# Step 8: Print summary
print("\n========== Summary ==========")
print(f"Total number of groups formed: {group_count}")
print(f"Total rows across all groups: {total_group_size}")
print(f"Original total rows in result_df: {len(result_df)}")
missing_rows = len(result_df) - total_group_size
print(f"Missing or uncounted rows (should be 0): {missing_rows}")

# Step 9: Export to Excel
output_file = "merged_ce_pe_variance_summary_debug.xlsx"
final_df.to_excel(output_file, index=False)
print(f"\nSaved debug summary to: {output_file}")
