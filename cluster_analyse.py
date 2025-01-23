import pandas as pd

# Paths to the normalized data file and output file
normalized_file = "C:\\Users\\Bastiaan\\Documents\\Uni\\scriptie\\scriptie_excl\\normalized_rivalry.xlsx"
output_file = "C:\\Users\\Bastiaan\\Documents\\Uni\\scriptie\\scriptie_excl\\weighted_rivalry_cluster_analysis_output.xlsx"

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

# Define clusters
clusters = {
    "Vergelijkbaarheid": ['inwonerrang_norm', 'dialect_norm', 'provincie_norm'],
    "Economisch": ['commerciele_sectoren_norm', 'toerisme_banen_norm', 'belangrijke_hubs_norm'],
    "Sociaal-Cultureel": ['web_cooccurrence_norm', 'voetbal_avg', 'steden_straal_norm']  # This cluster groups factors related to social or cultural dynamics.
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

# Initialize a results DataFrame
results_summary = []

# Perform cluster analysis
for cluster_name, cluster_factors in clusters.items():
    # Calculate weighted scores for the cluster
    weighted_results = data.copy()
    for col in cluster_factors:
        if col in data.columns:
            weighted_results[f"{col}_weighted"] = data[col] * weights[col]

    # Calculate total weighted score for the cluster
    weighted_results['total_weighted_score'] = weighted_results[
        [f"{col}_weighted" for col in cluster_factors]
    ].sum(axis=1)

    # Rank the cities based on the cluster's total weighted score
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

    # Save the result for this cluster
    results_summary.append({
        "cluster": cluster_name,
        "top_10": counts["top_10"],
        "top_25": counts["top_25"],
        "top_50": counts["top_50"],
        "top_100": counts["top_100"],
    })

# Convert the summary into a DataFrame
summary_df = pd.DataFrame(results_summary)

# Save the results to an Excel file
with pd.ExcelWriter(output_file) as writer:
    summary_df.to_excel(writer, sheet_name="Cluster_Analysis", index=False)

print(f"Cluster analysis saved to {output_file} in the 'Cluster_Analysis' sheet.")
