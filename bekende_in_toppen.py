import pandas as pd

# Path to the weighted results file
weighted_file_path = "C:\\Users\\Bastiaan\\Documents\\Uni\\scriptie\\scriptie_excl\\weighted_rivalry_output.xlsx"

# Load the datasets
normalized_data = pd.read_excel(weighted_file_path, sheet_name='Normalized_Data')  # Normal scores
weighted_data = pd.read_excel(weighted_file_path, sheet_name='Weighted_Results')  # Weighted scores

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

# Normalize city names by stripping extra spaces and converting to lowercase
def capitalize_city_names(city):
    if city == "den haag":
        return "Den Haag"
    elif city == "den bosch":
        return "Den Bosch"
    else:
        return city.title()

normalized_data['city_a'] = normalized_data['city_a'].str.strip().str.lower().apply(capitalize_city_names)
normalized_data['city_b'] = normalized_data['city_b'].str.strip().str.lower().apply(capitalize_city_names)
weighted_data['city_a'] = weighted_data['city_a'].str.strip().str.lower().apply(capitalize_city_names)
weighted_data['city_b'] = weighted_data['city_b'].str.strip().str.lower().apply(capitalize_city_names)

# Sort both datasets by their respective scores (normal and weighted)
normalized_data_sorted = normalized_data.sort_values(by="total_rivalry_score", ascending=False).reset_index(drop=True)
weighted_data_sorted = weighted_data.sort_values(by="total_weighted_score", ascending=False).reset_index(drop=True)

# Add rank columns for both normal and weighted scores
normalized_data_sorted['rank'] = normalized_data_sorted.index + 1
weighted_data_sorted['rank'] = weighted_data_sorted.index + 1

# Function to retrieve the position of a city pair (or its reverse) in the sorted data
def get_pair_rank(city_a, city_b, ranked_df, score_column):
    match = ranked_df[
        ((ranked_df['city_a'] == city_a) & (ranked_df['city_b'] == city_b)) |
        ((ranked_df['city_a'] == city_b) & (ranked_df['city_b'] == city_a))
    ]
    if not match.empty:
        return match.iloc[0][score_column]
    else:
        return "Not Found"

# Create an empty list to store the results for each city pair
rivalen_data = []

# Iterate over the list of city pairs and retrieve their rankings
for city_a, city_b in city_pairs:
    normal_position = get_pair_rank(city_a, city_b, normalized_data_sorted, 'rank')
    weighted_position = get_pair_rank(city_a, city_b, weighted_data_sorted, 'rank')
    rivalen_data.append({
        'city_a': city_a,
        'city_b': city_b,
        'normal_position': normal_position,
        'weighted_position': weighted_position
    })

# Convert the list of results to a DataFrame
rivalen_data = pd.DataFrame(rivalen_data)

# Calculate counts for the top 10, top 25, top 50, and top 100 positions based on the normal and weighted ranking
top_10_count_normal = rivalen_data[rivalen_data['normal_position'].apply(lambda x: isinstance(x, int) and x <= 10)].shape[0]
top_25_count_normal = rivalen_data[rivalen_data['normal_position'].apply(lambda x: isinstance(x, int) and x <= 25)].shape[0]
top_50_count_normal = rivalen_data[rivalen_data['normal_position'].apply(lambda x: isinstance(x, int) and x <= 50)].shape[0]
top_100_count_normal = rivalen_data[rivalen_data['normal_position'].apply(lambda x: isinstance(x, int) and x <= 100)].shape[0]

top_10_count_weighted = rivalen_data[rivalen_data['weighted_position'].apply(lambda x: isinstance(x, int) and x <= 10)].shape[0]
top_25_count_weighted = rivalen_data[rivalen_data['weighted_position'].apply(lambda x: isinstance(x, int) and x <= 25)].shape[0]
top_50_count_weighted = rivalen_data[rivalen_data['weighted_position'].apply(lambda x: isinstance(x, int) and x <= 50)].shape[0]
top_100_count_weighted = rivalen_data[rivalen_data['weighted_position'].apply(lambda x: isinstance(x, int) and x <= 100)].shape[0]

# Create a summary DataFrame for the counts
summary_data = pd.DataFrame({
    'Top Rank': ['Top 10', 'Top 25', 'Top 50', 'Top 100'],
    'Aantal (normaal)': [top_10_count_normal, top_25_count_normal, top_50_count_normal, top_100_count_normal],
    'Aantal (gewogen)': [top_10_count_weighted, top_25_count_weighted, top_50_count_weighted, top_100_count_weighted]
})

# Save the summary and the updated rivalen_data to a new Excel file
output_file_path = "C:\\Users\\Bastiaan\\Documents\\Uni\\scriptie\\scriptie_excl\\updated_rivalen_positions.xlsx"

with pd.ExcelWriter(output_file_path) as writer:
    # Write the summary to the first sheet
    summary_data.to_excel(writer, sheet_name='Summary', index=False)
    # Write the updated rivalen_data to a second sheet
    rivalen_data.to_excel(writer, sheet_name='Rivalen Positions', index=False)

print(f"Updated positions and summary saved to {output_file_path}")
