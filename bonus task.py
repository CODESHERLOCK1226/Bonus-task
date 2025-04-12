import pandas as pd
import numpy as np

# URLs for the spreadsheets
url1 = 'https://docs.google.com/spreadsheets/d/1PIjSBRUy8rNJoChCVUa7ElYaac5lPFWF/export?format=xlsx'
url2 = 'https://docs.google.com/spreadsheets/d/1_eZoQk462qtDs6jas9nTx6b7kkJiU_zv/export?format=xlsx'

def add_requested_columns(df):
    """Add columns for Participant Name, Team Name, Participant Type, Participant Phone Number, and Participant Email."""
    
    # Add missing columns with blank values
    for col in ['Participant Name', 'Team Name', 'Participant Type', 'Participant Phone Number', 'Participant Email']:
        if col not in df.columns:
            df[col] = pd.NA
    
    return df

def clean_data(df):
    """Handle missing data and ensure consistency."""
    # Fill missing phone numbers with 'N/A'
    df['Participant Phone Number'] = df['Participant Phone Number'].fillna('N/A')
    
    # Normalize participant type to lowercase
    if 'Participant Type' in df.columns:
        df['Participant Type'] = df['Participant Type'].str.lower().str.strip()
    
    # Replace missing critical fields with placeholder values
    critical_fields = ['Participant Name', 'Participant Email']
    for field in critical_fields:
        if field in df.columns:
            df[field] = df[field].fillna('Unknown')
    
    return df

# Load both spreadsheets
try:
    df1 = pd.read_excel(url1)
    df2 = pd.read_excel(url2)
except Exception as e:
    print(f"Error loading spreadsheets: {e}")
else:
    print("Spreadsheets loaded successfully.")
    
    # Add requested columns
    df1 = add_requested_columns(df1)
    df2 = add_requested_columns(df2)
    
    # Combine datasets
    combined_df = pd.concat([df1, df2], ignore_index=True)
    
    # Clean and validate data
    combined_df = clean_data(combined_df)
    
    output_file = 'combined_data.xlsx'
    
    try:
        combined_df.to_excel(output_file, index=False)
        print(f"Data saved to {output_file}.")
    except Exception as e:
        print(f"Failed to save data: {e}")
