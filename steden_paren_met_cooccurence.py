import pandas as pd
import itertools
import math
import numpy as np  # Voor NaN-waarden

# Function to calculate haversine distance between two cities using latitude and longitude
def haversine(lat1, lon1, lat2, lon2):
    R = 6371.0  # Radius of the Earth in kilometers
    lat1_rad, lon1_rad, lat2_rad, lon2_rad = map(math.radians, [lat1, lon1, lat2, lon2])
    dlat = lat2_rad - lat1_rad
    dlon = lon2_rad - lon1_rad
    a = math.sin(dlat / 2)**2 + math.cos(lat1_rad) * math.cos(lat2_rad) * math.sin(dlon / 2)**2
    c = 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))
    return R * c

# Dialect groups mapping
dialect_groups = {
    "A": [1, 2], "B": [3, 4, 5, 6, 7, 8, 9], "C": [10, 11, 12, 13, 14, 15, 16, 17, 18],
    "D": [19], "E": [20, 21, 22, 23], "F": [24]
}

# Voetbal derby pairs
voetbal_derbys = {
    ("Rotterdam", "Amsterdam"), ("Amsterdam", "Eindhoven"), ("Eindhoven", "Rotterdam"),
    ("Tilburg", "Breda"), ("Oss", "Den Bosch"), ("Eindhoven", "Helmond"),
    ("Groningen", "Heerenveen"), ("Groningen", "Emmen"), ("Rotterdam", "Utrecht"),
    ("Amsterdam", "Utrecht"), ("Eindhoven", "Utrecht"), ("Heerenveen", "Leeuwarden"),
    ("Amsterdam", "Alkmaar"), ("Dordrecht", "Rotterdam"), ("Rotterdam", "Den Haag"),
    ("Arnhem", "Nijmegen"), ("Amsterdam", "Den Haag"), ("Zwolle", "Deventer"),
    ("Zwolle", "Enschede"), ("Zwolle", "Almelo"), ("Enschede", "Deventer"),
    ("Enschede", "Almelo"), ("Deventer", "Almelo"), ("Velsen", "Amsterdam"),
    ("Velsen", "Alkmaar")
}

# Pas hier je eigen paths toe voor bestanden
cities_file = "steden_data"
output_file = "stad_data_output"
additional_input_file = "web_co-occurence_file"
rivalry_matrix_file = "steden_paren_data"
filtered_output_file = "filtered_data"

try:
    # Step 1: Read the cities data
    df_cities = pd.read_excel(cities_file)
    df_cities.columns = df_cities.columns.str.strip()
    necessary_columns = [
        'stad', 'inwonerrang', 'Latitude', 'Longitude', 'Commerciële_sectoren_nu',
        'Toerisme_banen', 'Provincie', 'Dialect', 'Historische_rangorde_Eredivisie'
    ]
    for col in necessary_columns:
        if col not in df_cities.columns:
            raise KeyError(f"'{col}' column not found in the cities dataset")

    # Generate all unique pairs of cities
    cities_list = df_cities['stad'].tolist()
    city_pairs = list(itertools.combinations(cities_list, 2))
    data = []

    for city_a, city_b in city_pairs:
        city_a_data = df_cities[df_cities['stad'] == city_a].iloc[0]
        city_b_data = df_cities[df_cities['stad'] == city_b].iloc[0]

        inwonerrang_diff = abs(city_a_data['inwonerrang'] - city_b_data['inwonerrang'])
        afstand = haversine(city_a_data['Latitude'], city_a_data['Longitude'],
                            city_b_data['Latitude'], city_b_data['Longitude'])

        city_a_sectors = set(city_a_data['Commerciële_sectoren_nu'].split(','))
        city_b_sectors = set(city_b_data['Commerciële_sectoren_nu'].split(','))
        common_sectors = len(city_a_sectors & city_b_sectors)

        toerisme_banen_diff = abs(city_a_data['Toerisme_banen'] - city_b_data['Toerisme_banen'])
        hubs_columns = ['Vliegveld', 'Haven', 'Academisch_ziekenhuis', 'Grote_evenementhal']
        belangrijke_hubs_score = sum(
            city_a_data.get(hub, 0) + city_b_data.get(hub, 0) == 2 for hub in hubs_columns
        )

        nearby_cities_count = sum(
            haversine(city_a_data['Latitude'], city_a_data['Longitude'], city['Latitude'], city['Longitude']) <= 25
            or haversine(city_b_data['Latitude'], city_b_data['Longitude'], city['Latitude'], city['Longitude']) <= 25
            for _, city in df_cities.iterrows() if city['stad'] not in {city_a, city_b}
        )

        dezelfde_provincie = city_a_data['Provincie'] == city_b_data['Provincie']
        city_a_dialect = city_a_data['Dialect']
        city_b_dialect = city_b_data['Dialect']
        dialect_group_a = next((g for g, members in dialect_groups.items() if city_a_dialect in members), None)
        dialect_group_b = next((g for g, members in dialect_groups.items() if city_b_dialect in members), None)
        if dialect_group_a and dialect_group_b:
            dialect_relation = 'strong' if dialect_group_a == dialect_group_b else 'average'
        else:
            dialect_relation = 'none'

        voetbal_derby = (city_a, city_b) in voetbal_derbys or (city_b, city_a) in voetbal_derbys
        rank_a = city_a_data['Historische_rangorde_Eredivisie']
        rank_b = city_b_data['Historische_rangorde_Eredivisie']
        historische_rangorde_diff = np.nan if rank_a == 0 or rank_b == 0 else abs(rank_a - rank_b)

        row = {
            'city_a': city_a, 'city_b': city_b, 'inwonerrang': inwonerrang_diff,
            'afstand': afstand, 'commerciële_sectoren': common_sectors,
            'toerisme_banen': toerisme_banen_diff, 'belangrijke_hubs': belangrijke_hubs_score,
            'steden_straal_25km': nearby_cities_count, 'dezelfde_provincie': dezelfde_provincie,
            'dialect': dialect_relation, 'voetbal_derby': voetbal_derby,
            'historische_rangorde_eredivisie': historische_rangorde_diff
        }
        data.append(row)

    # Save to stad_data_output.xlsx
    df_output = pd.DataFrame(data)
    df_output.to_excel(output_file, index=False)

    # Step 2: Process additional data
    df_main = pd.read_excel(additional_input_file)
    df_rivalry = pd.read_excel(rivalry_matrix_file)
    relevant_data = df_main[["city_a", "city_b", "web_co-occurence"]]
    relevant_data_reversed = relevant_data.rename(columns={"city_a": "city_b", "city_b": "city_a"})
    combined_data = pd.concat([relevant_data, relevant_data_reversed], ignore_index=True)
    combined_data_grouped = combined_data.groupby(["city_a", "city_b"], as_index=False)["web_co-occurence"].mean()
    filtered_data = pd.merge(df_rivalry, combined_data_grouped, on=["city_a", "city_b"], how="left")
    filtered_data = filtered_data[["city_a", "city_b", "web_co-occurence"]]
    filtered_data.to_excel(filtered_output_file, index=False)

    print(f"Data processing complete. Output saved to:\n- {output_file}\n- {filtered_output_file}")

except Exception as e:
    print(f"An error occurred: {e}")