import pandas as pd

# Paths to the normalized data file and output file
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

# List of city pairs to analyze
city_pairs = [
    ('Amsterdam', 'Rotterdam'),
    ('Eindhoven', 'Tilburg'),
    ('Rotterdam', 'Den Haag'),
    ('Breda', 'Tilburg'),
    ('Utrecht', 'Amsterdam'),
    ('Den Bosch', 'Tilburg'),
    ('Oss', 'Den Bosch'),
    ('Arnhem', 'Nijmegen'),
    ('Zwolle', 'Deventer'),
    ('Schiedam', 'Vlaardingen'),
]

# Generate both original and reversed city pairs for matching
city_pairs_set = set(city_pairs + [(b, a) for a, b in city_pairs])

# Factors for weight adjustments
adjustments = [0.9, 0.8, 1.1, 1.2]  # 10% decrease, 20% decrease, 10% increase, 20% increase

# Initialize a results DataFrame
results_summary = []

# Calculate positions without any weight adjustments (baseline)
weighted_results = data.copy()
for col, weight in weights.items():
    if col in data.columns:
        weighted_results[f"{col}_weighted"] = data[col] * weight

# Calculate total weighted score
weighted_results['total_weighted_score'] = weighted_results[
    [col for col in weighted_results.columns if '_weighted' in col]
].sum(axis=1)

# Rank the cities based on total weighted score
weighted_results['rank'] = weighted_results['total_weighted_score'].rank(ascending=False, method='min')

# Analyze the positions of the specific city pairs for the baseline
counts = {"top_10": 0, "top_25": 0, "top_50": 0, "top_100": 0}
for _, row in weighted_results.iterrows():
    pair = (row['city_a'], row['city_b'])
    if pair in city_pairs_set:
        rank = row['rank']
        if rank <= 10:
            counts["top_10"] += 1
        if rank <= 25:
            counts["top_25"] += 1
        if rank <= 50:
            counts["top_50"] += 1
        if rank <= 100:
            counts["top_100"] += 1

# Save the baseline result
results_summary.append({
    "factor": "baseline",
    "adjustment": 1.0,
    "top_10": counts["top_10"],
    "top_25": counts["top_25"],
    "top_50": counts["top_50"],
    "top_100": counts["top_100"],
})

# Iterate over each factor and apply weight adjustments
for factor, original_weight in weights.items():
    for adjustment in adjustments:
        # Create adjusted weights
        adjusted_weights = weights.copy()
        adjusted_weights[factor] = original_weight * adjustment

        # Calculate weighted scores
        weighted_results = data.copy()
        for col, weight in adjusted_weights.items():
            if col in data.columns:
                weighted_results[f"{col}_weighted"] = data[col] * weight

        # Calculate total weighted score
        weighted_results['total_weighted_score'] = weighted_results[
            [col for col in weighted_results.columns if '_weighted' in col]
        ].sum(axis=1)

        # Rank the cities based on total weighted score
        weighted_results['rank'] = weighted_results['total_weighted_score'].rank(ascending=False, method='min')

        # Analyze the positions of the specific city pairs
        counts = {"top_10": 0, "top_25": 0, "top_50": 0, "top_100": 0}
        for _, row in weighted_results.iterrows():
            pair = (row['city_a'], row['city_b'])
            if pair in city_pairs_set:
                rank = row['rank']
                if rank <= 10:
                    counts["top_10"] += 1
                if rank <= 25:
                    counts["top_25"] += 1
                if rank <= 50:
                    counts["top_50"] += 1
                if rank <= 100:
                    counts["top_100"] += 1

        # Save the result for this adjustment
        results_summary.append({
            "factor": factor,
            "adjustment": adjustment,
            "top_10": counts["top_10"],
            "top_25": counts["top_25"],
            "top_50": counts["top_50"],
            "top_100": counts["top_100"],
        })

# Convert the summary into a DataFrame
summary_df = pd.DataFrame(results_summary)

# Save the results to an Excel file
with pd.ExcelWriter(output_file) as writer:
    summary_df.to_excel(writer, sheet_name="Weight_Analysis", index=False)

print(f"Weight analysis saved to {output_file} in the 'Weight_Analysis' sheet.")
