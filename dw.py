import pandas as pd

# ----- STEP 1: Helper - categorize limit -----
def categorize_limit(limit):
    try:
        val = float(str(limit).replace('K', '').replace(',', '').strip())
        if val < 50:
            return '<50K'
        elif 50 <= val <= 80:
            return '50K-80K'
        else:
            return '>80K'
    except:
        return 'Unknown'

# ----- STEP 2: Helper - map final_rating -----
def map_rating(rating):
    rating_str = str(rating).strip().upper()
    if rating_str.startswith('A'):
        return 'A'
    elif rating_str.startswith('B'):
        return 'B'
    else:
        return 'Other'

# ----- STEP 3: Prepare dataframe -----
result_df['limit_category'] = result_df['limit'].apply(categorize_limit)
result_df['rating_category'] = result_df['final_rating'].apply(map_rating)

# ----- STEP 4: Define -----
group_cols = ['date', 'risk_profile', 'limit_category', 'final_product', 'rating_category']
limit_order = ['<50K', '50K-80K', '>80K']
thresholds = list(range(5, 105, 5))  # 5 to 100

# ----- STEP 5: Compute counts -----
final_rows = []

for date in result_df['date'].dropna().unique():
    sub_df = result_df[result_df['date'] == date]
    grouped = sub_df.groupby(group_cols)

    for group_keys, group in grouped:
        row = dict(zip(group_cols, group_keys))

        # Count CE thresholds
        for t in thresholds:
            row[f'ce_{t}'] = group[group['ce_variance'] > t].shape[0]

        # Count PE thresholds
        for t in thresholds:
            row[f'pe_{t}'] = group[group['pe_variance'] > t].shape[0]

        final_rows.append(row)

# ----- STEP 6: Create DataFrame and format -----
final_df = pd.DataFrame(final_rows)

# Optional: Sort & order columns
final_df['limit_category'] = pd.Categorical(final_df['limit_category'], categories=limit_order, ordered=True)
final_df = final_df.sort_values(['date', 'risk_profile', 'limit_category', 'final_product', 'rating_category'])

# Column ordering
ce_cols = [f'ce_{t}' for t in thresholds]
pe_cols = [f'pe_{t}' for t in thresholds]
ordered_cols = group_cols + ce_cols + pe_cols
final_df = final_df[ordered_cols]

# ----- STEP 7: Save result -----
output_file = "merged_ce_pe_variance_summary.xlsx"
final_df.to_excel(output_file, index=False)
print(f"Saved merged summary to: {output_file}")
