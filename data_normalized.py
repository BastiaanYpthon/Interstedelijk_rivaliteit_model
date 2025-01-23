import pandas as pd

# Path to the input Excel file
input_file = "C:\\Users\\Bastiaan\\Documents\\Uni\\scriptie\\scriptie_excl\\stad_data_output.xlsx"
output_file = "C:\\Users\\Bastiaan\\Documents\\Uni\\scriptie\\scriptie_excl\\normalized_rivalry.xlsx"

# Load the dataset
data = pd.read_excel(input_file)

# Helper functions for normalization
def normalize_inwonerrang(diff):
    return 1 - (diff - 1) / (57 - 1) if not pd.isna(diff) else 0

def normalize_afstand(afstand, min_afstand, max_afstand):
    if max_afstand == min_afstand:
        return 1  # Als alle waarden gelijk zijn
    return 1 - (afstand - min_afstand) / (max_afstand - min_afstand)

def normalize_sectoren(count):
    return {5: 1, 4: 0.8, 3: 0.6, 2: 0.4, 1: 0.2, 0: 0}.get(count, 0)

def normalize_toerisme_banen(diff):
    return max(1 - (diff - 10) / (80510 - 10), 0) if not pd.isna(diff) else 0

def normalize_hubs(count):
    return {4: 1, 3: 0.75, 2: 0.5, 1: 0.25, 0: 0}.get(count, 0)

def normalize_steden_straal(count):
    if pd.isna(count):
        return 0
    return max(-1 + (count / 16), -1)

def normalize_dialect(value):
    return {"strong": 1, "average": 0.5, "none": 0}.get(value, 0)

def normalize_web_cooccurrence(value, min_value):
    if pd.isna(value) or value >= 0:
        return 0  # Waarden groter dan of gelijk aan 0 krijgen waarde 0
    return -1 + (value - min_value) / (0 - min_value)

def normalize_province(same_province):
    return 1 if same_province else 0

def normalize_eredivisie_difference(diff, max_rank=16):
    # Als er geen data is (NaN), return 0
    if pd.isna(diff):
        return 0
    
    # Als het verschil gelijk is aan het minimum (beste verschil), return 1
    if diff == 1:
        return 1
    
    # Als het verschil gelijk is aan het maximum (slechtste verschil), return 0
    if diff == max_rank:
        return 0

    # Genormaliseerde waarde berekenen als lineaire schaal van 1 naar 0
    max_difference = max_rank - 1  # Grootste mogelijke verschil (van 1 naar max_rank)
    normalized_value = 1 - (diff - 1) / max_difference  # Normalisatie van 1 naar 0

    # Zorg ervoor dat de waarde niet groter dan 1 of kleiner dan 0 is
    return max(0, min(normalized_value, 1))

def normalize_football_derby(is_derby):
    return 1 if is_derby else 0

# Calculate min and max values for distance and web co-occurrence
min_afstand = data['afstand'].min()
max_afstand = data['afstand'].max()
min_web_cooccurrence = data['web_co-occurrence'].min()
max_web_cooccurrence = data['web_co-occurrence'].max()

# Normalize columns
data['inwonerrang_norm'] = data['inwonerrang'].apply(normalize_inwonerrang)
data['afstand_norm'] = data['afstand'].apply(normalize_afstand, args=(min_afstand, max_afstand))
data['commerciele_sectoren_norm'] = data['commerciële_sectoren'].apply(normalize_sectoren)
data['toerisme_banen_norm'] = data['toerisme_banen'].apply(normalize_toerisme_banen)
data['belangrijke_hubs_norm'] = data['belangrijke_hubs'].apply(normalize_hubs)
data['steden_straal_norm'] = data['steden_straal_25km'].apply(normalize_steden_straal)
data['dialect_norm'] = data['dialect'].apply(normalize_dialect)
data['web_cooccurrence_norm'] = data['web_co-occurrence'].apply(
    normalize_web_cooccurrence, args=(min_web_cooccurrence,)
)
data['provincie_norm'] = data['dezelfde_provincie'].apply(normalize_province)
data['eredivisie_norm'] = data['historische_rangorde_eredivisie'].apply(normalize_eredivisie_difference)
data['voetbal_derby_norm'] = data['voetbal_derby'].apply(normalize_football_derby)

# Calculate football score as the average of the last two columns
data['voetbal_avg'] = data[['eredivisie_norm', 'voetbal_derby_norm']].mean(axis=1)

# Calculate total rivalry score
data['total_rivalry_score'] = data[['inwonerrang_norm', 'afstand_norm', 'commerciele_sectoren_norm',
                                    'toerisme_banen_norm', 'belangrijke_hubs_norm', 'steden_straal_norm',
                                    'dialect_norm', 'web_cooccurrence_norm', 'provincie_norm',
                                    'voetbal_avg']].sum(axis=1)

# Insert a blank column as a separator
data.insert(len(data.columns) // 2, '---', '')

# Save the normalized data to a new Excel file
data.to_excel(output_file, index=False)

print(f"Normalized data saved to {output_file}")