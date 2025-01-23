import pandas as pd

# Path to the normalized data file
normalized_file = "C:\\Users\\Bastiaan\\Documents\\Uni\\scriptie\\scriptie_excl\\normalized_rivalry.xlsx"
output_file = "C:\\Users\\Bastiaan\\Documents\\Uni\\scriptie\\scriptie_excl\\weighted_rivalry_output.xlsx"

# Load the normalized data
data = pd.read_excel(normalized_file)

# Define weights for the factors
weights = {
    'inwonerrang_norm': 3.67,
    'afstand_norm': 3.79,
    'commerciele_sectoren_norm': 3.19,
    'toerisme_banen_norm': 2.48,
    'belangrijke_hubs_norm': 3.40,
    'steden_straal_norm': 2.93,
    'dialect_norm': 2.71,
    'web_cooccurrence_norm': 2.69,
    'provincie_norm': 3.31,
    'voetbal_avg': 4.31,
}

# Create a new DataFrame for weighted results
weighted_results = pd.DataFrame()

# Copy city columns to the weighted results
weighted_results['city_a'] = data['city_a']
weighted_results['city_b'] = data['city_b']

# Apply weights and calculate weighted scores for each factor
for factor, weight in weights.items():
    if factor in data.columns:
        weighted_results[f"{factor}_weighted"] = data[factor] * weight

# Calculate the total weighted score
weighted_results['total_weighted_score'] = weighted_results[
    [col for col in weighted_results.columns if '_weighted' in col]
].sum(axis=1)

# Calculate the average weight
average_weight = sum(weights.values()) / len(weights)

# Divide the total weighted score by the average weight
weighted_results['adjusted_weighted_score'] = weighted_results['total_weighted_score'] / average_weight

# Save the weighted results to a new Excel file
with pd.ExcelWriter(output_file) as writer:
    data.to_excel(writer, sheet_name="Normalized_Data", index=False)
    weighted_results.to_excel(writer, sheet_name="Weighted_Results", index=False)

print(f"Weighted results saved to {output_file} in the 'Weighted_Results' sheet.")