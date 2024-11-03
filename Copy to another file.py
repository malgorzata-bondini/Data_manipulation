import pandas as pd
from dateutil import parser

def copy_and_transform_data(pum_path, manual_path, pum_sheet_name='2324', manual_sheet_name='Price Update Template'):
    # Define the columns
    columns_mapping = {
        'Supplier Name': 'Supplier Name',
        'Supplier Number': 'Supplier Number',
        'Item Number': 'Item Number',
        'New price': 'New price',
        'Item Price Currency': 'Item Price Currency',
        'Branch Plant': 'Branch Plant',
        'Unit of Measures': 'Unit of measure',
        'Main Reason 1': 'Main Reason 1',
        'Weight (%) Reason 1': 'Weight (%) Reason 1',
        'Main Reason 2': 'Main Reason 2',
        'Weight (%) Reason 2': 'Weight (%) Reason 2',
        'Main Reason 3': 'Main Reason 3',
        'Weight (%) Reason 3': 'Weight (%) Reason 3',
        'Out-going Item number': 'Out-going Item number',
        'Quantity break': 'Quantity Breaks',
        'CP Project name': 'Project name',
        'Effective date': 'Effective date (DDMMYYYY)',
        'Update date': 'Update date',
        'Item responsible': 'Requester',
        'CM05 BP': 'CM05 BP',
        'CM05': 'CM05'
    }

    # Read the data
    try:
        pum_df = pd.read_excel(pum_path, sheet_name=pum_sheet_name, header=1, dtype=str)
        if pum_df.empty:
            print("PUM DataFrame is empty.")
            return
    except Exception as e:
        raise ValueError(f"Error reading PUM file: {e}")
    
    # Trim
    pum_df = pum_df.apply(lambda x: x.str.strip() if x.dtype == "object" else x)

    # Empty or NaN values
    if pum_df['Item Number'].dropna().empty:
        print("No items to process in the PUM DataFrame.")
        return
    
    # Read the new file
    try:
        manual_df = pd.read_excel(manual_path, sheet_name=manual_sheet_name, dtype=str)
        if manual_df.empty:
            print("Manual DataFrame is empty. Initializing new DataFrame.")
            manual_df = pd.DataFrame(columns=columns_mapping.values())
    except Exception as e:
        raise ValueError(f"Error reading Manual file: {e}")
    for manual_col in columns_mapping.values():
        if manual_col not in manual_df.columns:
            manual_df[manual_col] = ''

    # Copy and transform data
    for pum_col, manual_col in columns_mapping.items():
        if pum_col in pum_df.columns:
            manual_df[manual_col] = pum_df[pum_col]
        else:
            print(f"The column '{pum_col}' does not exist in the PUM DataFrame.")
            manual_df[manual_col] = ''
    manual_df['Supp No'] = '620'
    manual_df['Branch Plant'] = manual_df['Branch Plant'].astype(str)

    # Calculatations
    def calculate_currency(branch_plant):
        if pd.isna(branch_plant):
            return None
        if branch_plant.startswith('490'):
            return 'HUF'
        elif branch_plant.startswith('110'):
            return 'DKK'
        elif branch_plant.startswith('620'):
            return 'CNY'
        elif branch_plant.startswith('711') or branch_plant.startswith('710'):
            return 'USD'
        elif branch_plant.startswith('790'):
            return 'CRC'
        elif branch_plant.startswith('351'):
            return 'EUR'
        else:
            return None

    manual_df['CM05 BP Currency'] = manual_df['Branch Plant'].apply(calculate_currency)

    # Convert date
    def format_date(date_str):
        if pd.isna(date_str):
            return date_str
        try:
            return parser.parse(date_str, dayfirst=True).strftime('%d.%m.%Y')
        except Exception as e:
            print(f"Error parsing date '{date_str}': {e}")
            return date_str

    manual_df['Effective date (DDMMYYYY)'] = manual_df['Effective date (DDMMYYYY)'].apply(format_date)
    manual_df['Update date'] = manual_df['Update date'].apply(format_date)

    # Replace dots with commas
    manual_df['New price'] = manual_df['New price'].str.replace('.', ',', regex=False)
    manual_df['CM05'] = manual_df['CM05'].str.replace('.', ',', regex=False)

    # Save
    try:
        with pd.ExcelWriter(manual_path, engine='openpyxl', mode='w') as writer:
            manual_df.to_excel(writer, sheet_name=manual_sheet_name, index=False)
        print("The file was updated successfully")
    except Exception as e:
        raise ValueError(f"Error saving Manual file: {e}")

# Paths
pum_path = 'C:\\Users\\plmala\\OneDrive - Coloplast A S\\Desktop\\Python\\PPM\\Price Updates (Python)\\PUM\\PUM.xlsx'
manual_path = 'C:\\Users\\plmala\\OneDrive - Coloplast A S\\Desktop\\Python\\PPM\\Price Updates (Python)\\PUM\\Manual.xlsx'

copy_and_transform_data(pum_path, manual_path)
