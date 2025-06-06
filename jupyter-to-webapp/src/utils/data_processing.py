def clean_data(df):
    """Cleans the input DataFrame by removing NaN values and resetting the index."""
    df_cleaned = df.dropna().reset_index(drop=True)
    return df_cleaned

def prepare_data(df):
    """Prepares the input DataFrame for analysis by converting data types and filtering."""
    # Example: Convert flow column to numeric
    df['Q [m³/h]'] = pd.to_numeric(df['Q [m³/h]'], errors='coerce')
    # Filter out negative values
    df = df[df['Q [m³/h]'] >= 0]
    return df

def extract_curve_data(df, curve_type):
    """Extracts specific curve data based on the curve type."""
    if curve_type not in ['Efficiency', 'Power', 'Head', 'NPSH']:
        raise ValueError(f"Invalid curve type: {curve_type}")
    
    return df[['Q [m³/h]', f'{curve_type} [metric]']]  # Adjust column name as needed

def apply_affinity_laws(df, n1, n2):
    """Applies affinity laws to the DataFrame based on the given speeds."""
    df['Q [m³/h]'] *= (n2 / n1)
    df['H [m]'] *= (n2 / n1) ** 2
    df['P2 [kW]'] *= (n2 / n1) ** 3
    df['NPSH [m]'] *= (n2 / n1) ** 2
    return df