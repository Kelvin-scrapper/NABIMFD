import pandas as pd
import os
import re
from datetime import datetime

# --- Constants ---
INPUT_DIR = 'downloads'
OUTPUT_DIR = 'output'
SOURCE_FILE = 'BORROWINGS.xls'  # Actually a TSV file
OUTPUT_FILE = 'NABIMFD_OUTPUT.xlsx'

COUNTRY_CODES = {
    'Australia': 'NABIMFD.AUS.M',
    'Austria': 'NABIMFD.AUT.M',
    'Belgium': 'NABIMFD.BEL.M',
    'Brazil': 'NABIMFD.BRA.M',
    'Canada': 'NABIMFD.CAN.M',
    'Chile': 'NABIMFD.CHL.M',
    'China': 'NABIMFD.CHN.M',
    'Cyprus': 'NABIMFD.CYP.M',
    'Denmark': 'NABIMFD.DNK.M',
    'Finland': 'NABIMFD.FIN.M',
    'France': 'NABIMFD.FRA.M',
    'Germany': 'NABIMFD.DEU.M',
    'Greece': 'NABIMFD.GRC.M',
    'India': 'NABIMFD.IND.M',
    'Ireland': 'NABIMFD.IRL.M',
    'Israel': 'NABIMFD.ISR.M',
    'Italy': 'NABIMFD.ITA.M',
    'Japan': 'NABIMFD.JPN.M',
    'Korea': 'NABIMFD.KOR.M',
    'Kuwait': 'NABIMFD.KWT.M',
    'Luxembourg': 'NABIMFD.LUX.M',
    'Malaysia': 'NABIMFD.MYS.M',
    'Mexico': 'NABIMFD.MEX.M',
    'Netherlands': 'NABIMFD.NLD.M',
    'New Zealand': 'NABIMFD.NZL.M',
    'Norway': 'NABIMFD.NOR.M',
    'Philippines': 'NABIMFD.PHL.M',
    'Poland': 'NABIMFD.POL.M',
    'Portugal': 'NABIMFD.PRT.M',
    'Russian Federation': 'NABIMFD.RUS.M',
    'Russia': 'NABIMFD.RUS.M',
    'Saudi Arabia': 'NABIMFD.SAU.M',
    'Singapore': 'NABIMFD.SGP.M',
    'South Africa': 'NABIMFD.ZAF.M',
    'Spain': 'NABIMFD.ESP.M',
    'Sweden': 'NABIMFD.SWE.M',
    'Switzerland': 'NABIMFD.CHE.M',
    'Thailand': 'NABIMFD.THA.M',
    'United Kingdom': 'NABIMFD.GBR.M',
    'United States': 'NABIMFD.USA.M',
    'Hong Kong': 'NABIMFD.HKG.M'
}

STANDARD_COLUMN_ORDER = [
    '',
    'NABIMFD.AUS.M', 'NABIMFD.AUT.M', 'NABIMFD.BEL.M', 'NABIMFD.BRA.M',
    'NABIMFD.CAN.M', 'NABIMFD.CHL.M', 'NABIMFD.CHN.M', 'NABIMFD.CYP.M',
    'NABIMFD.DNK.M', 'NABIMFD.FIN.M', 'NABIMFD.FRA.M', 'NABIMFD.DEU.M',
    'NABIMFD.GRC.M', 'NABIMFD.IND.M', 'NABIMFD.IRL.M', 'NABIMFD.ISR.M',
    'NABIMFD.ITA.M', 'NABIMFD.JPN.M', 'NABIMFD.KOR.M', 'NABIMFD.KWT.M',
    'NABIMFD.LUX.M', 'NABIMFD.MYS.M', 'NABIMFD.MEX.M', 'NABIMFD.NLD.M',
    'NABIMFD.NZL.M', 'NABIMFD.NOR.M', 'NABIMFD.PHL.M', 'NABIMFD.POL.M',
    'NABIMFD.PRT.M', 'NABIMFD.RUS.M', 'NABIMFD.SAU.M', 'NABIMFD.SGP.M',
    'NABIMFD.ZAF.M', 'NABIMFD.ESP.M', 'NABIMFD.SWE.M', 'NABIMFD.CHE.M',
    'NABIMFD.THA.M', 'NABIMFD.GBR.M', 'NABIMFD.USA.M', 'NABIMFD.HKG.M'
]

def extract_date_from_tsv_file(file_path):
    """Extract the 'As of:' date from the TSV file and convert to YYYY-MM format"""
    try:
        # Read the file as text since it's a TSV file with .xls extension
        with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
            content = f.read()
        
        print(f"File content preview: {content[:500]}")
        
        # Look for "As of:" pattern with flexible spacing
        as_of_patterns = [
            r'As of:\s*([^\r\n\t]+)',
            r'As of\s*:\s*([^\r\n\t]+)',
            r'as of[:\s]+([^\r\n\t]+)',
        ]
        
        for pattern in as_of_patterns:
            as_of_match = re.search(pattern, content, re.IGNORECASE)
            if as_of_match:
                date_str = as_of_match.group(1).strip()
                print(f"Found date string with pattern '{pattern}': '{date_str}'")
                
                # Look for date patterns in the extracted string
                date_match = re.search(r'(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{1,2}),?\s+(\d{4})', date_str, re.IGNORECASE)
                if date_match:
                    month_name = date_match.group(1)
                    day = date_match.group(2)
                    year = date_match.group(3)
                    
                    # Convert to datetime and format as YYYY-MM
                    date_obj = datetime.strptime(f"{month_name} {day}, {year}", '%B %d, %Y')
                    formatted_date = date_obj.strftime('%Y-%m')
                    print(f"Extracted date: {month_name} {day}, {year} -> {formatted_date}")
                    return formatted_date
        
        # If no "As of:" found, look for any date pattern in the content
        print("'As of:' not found, searching for date patterns...")
        all_dates = re.findall(r'(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{1,2}),?\s+(\d{4})', content, re.IGNORECASE)
        
        if all_dates:
            # Use the first date found
            month_name, day, year = all_dates[0]
            date_obj = datetime.strptime(f"{month_name} {day}, {year}", '%B %d, %Y')
            formatted_date = date_obj.strftime('%Y-%m')
            print(f"Found date pattern: {month_name} {day}, {year} -> {formatted_date}")
            return formatted_date
        
        print("Could not find or parse any date, using current month")
        return datetime.now().strftime('%Y-%m')
        
    except Exception as e:
        print(f"Error extracting date: {e}")
        return datetime.now().strftime('%Y-%m')

def get_country_name_from_code(country_code):
    """Get country name from country code"""
    code_to_country = {v: k for k, v in COUNTRY_CODES.items()}
    return code_to_country.get(country_code, country_code.replace('NABIMFD.', '').replace('.M', ''))

def main():
    """Main function to run the data mapping process."""
    source_file_path = os.path.join(INPUT_DIR, SOURCE_FILE)

    if not os.path.exists(source_file_path):
        print(f"Error: Source file not found at '{source_file_path}'")
        return

    try:
        # 1. Extract the date from the TSV file
        report_date = extract_date_from_tsv_file(source_file_path)
        print(f"Using report date: {report_date}")

        # 2. Read the source file as a TSV (Tab-Separated Values)
        print("Reading TSV file...")
        columns = ['Member (Lender)', 'Member Code', 'Type', 'Facility', 'Start Date', 'Expiration of Term', 'Revolving', 'Amount', 'Currency', 'Amount Outstanding (SDR)', 'Status']
        
        # Read as TSV, skipping the header rows
        df = pd.read_csv(source_file_path, sep='\t', skiprows=8, names=columns, engine='python')
        
        # Clean the data - remove any rows that are obviously header information
        df = df.dropna(subset=['Member (Lender)', 'Facility', 'Amount'])
        df = df[df['Member (Lender)'].str.contains('Member', na=False) == False]  # Remove header rows
        
        print(f"Loaded {len(df)} rows of data")
        print("Sample data:")
        print(df.head())

        # 3. Filter for "New Arrangement to Borrow"
        facility_filter = "New Arrangement to Borrow"
        filtered_df = df[df['Facility'].str.strip() == facility_filter].copy()
        print(f"Found {len(filtered_df)} records for '{facility_filter}'")
        
        if len(filtered_df) == 0:
            print("No records found! Let's check what facilities are available:")
            print(df['Facility'].value_counts().head(10))

        # 4. Create the header rows and data row according to runbook format
        
        # First row: Timeseries IDs (CODEs)
        header_row_1 = STANDARD_COLUMN_ORDER.copy()
        
        # Second row: Timeseries descriptions
        header_row_2 = ['']  # Empty for first column
        for col in STANDARD_COLUMN_ORDER[1:]:  # Skip the first empty column
            country_name = get_country_name_from_code(col)
            header_row_2.append(f"New Arrangement to Borrow IMF Data: {country_name}")
        
        # Third row: Data values (start with the date in first column)
        data_row = [report_date]  # First column contains the date
        data_row.extend([None] * (len(STANDARD_COLUMN_ORDER) - 1))  # Fill rest with None
        
        # Enhanced country mapping with exact matches from IMF data
        enhanced_mapping = {
            'Australia': ['Australia', 'AUS'],
            'Austria': ['Austria', 'AUT'],
            'Belgium': ['Belgium', 'BEL'],
            'Brazil': ['Brazil', 'BRA'],
            'Canada': ['Canada', 'CAN'],
            'Chile': ['Chile', 'Banco Central de Chile', 'CHL'],
            'China': ['China', 'CHN'],
            'Cyprus': ['Cyprus', 'CYP'],
            'Denmark': ['Denmark', 'Danmarks Nationalbank', 'DNK'],
            'Finland': ['Finland', 'FIN'],
            'France': ['France', 'FRA'],
            'Germany': ['Germany', 'Deutsche Bundesbank', 'DEU'],
            'Greece': ['Greece', 'GRC'],
            'India': ['India', 'IND'],
            'Ireland': ['Ireland', 'IRL'],
            'Israel': ['Israel', 'Bank of Israel', 'ISR'],
            'Italy': ['Italy', 'ITA'],
            'Japan': ['Japan', 'JPN'],
            'Korea': ['Korea', 'KOR'],
            'Kuwait': ['Kuwait', 'KWT'],
            'Luxembourg': ['Luxembourg', 'LUX'],
            'Malaysia': ['Malaysia', 'MYS'],
            'Mexico': ['Mexico', 'MEX'],
            'Netherlands': ['Netherlands', 'The Netherlands', 'NLD'],
            'New Zealand': ['New Zealand', 'NZL'],
            'Norway': ['Norway', 'NOR'],
            'Philippines': ['Philippines', 'Bangko Sentral ng Pilipinas', 'PHL'],
            'Poland': ['Poland', 'Republic of', 'National Bank of Poland', 'POL'],
            'Portugal': ['Portugal', 'Banco de Portugal', 'PRT'],
            'Russian Federation': ['Russian Federation', 'Russia', 'RUS'],
            'Saudi Arabia': ['Saudi Arabia', 'SAU'],
            'Singapore': ['Singapore', 'SGP'],
            'South Africa': ['South Africa', 'ZAF'],
            'Spain': ['Spain', 'ESP'],
            'Sweden': ['Sweden', 'Sveriges Riksbank', 'SWE'],
            'Switzerland': ['Switzerland', 'Swiss National Bank', 'CHE'],
            'Thailand': ['Thailand', 'THA'],
            'United Kingdom': ['United Kingdom', 'GBR'],
            'United States': ['United States', 'USA'],
            'Hong Kong': ['Hong Kong', 'Hong Kong Monetary Authority', 'HKMA', 'HKG']
        }
        
        # Fill in the data values with PRECISE matching to avoid cross-mapping
        for _, row in filtered_df.iterrows():
            member_name = str(row['Member (Lender)']).strip()
            member_code = str(row['Member Code']).strip() if pd.notna(row['Member Code']) else ''
            amount_str = str(row['Amount']).strip()
            
            # Clean and parse amount
            try:
                # Remove commas and convert to float
                clean_amount = amount_str.replace(',', '').replace(' ', '')
                amount = float(clean_amount) if clean_amount and clean_amount != 'nan' else None
            except:
                amount = None
                
            if amount is not None:
                # PRECISE matching - prioritize exact matches first
                matched = False
                
                # First, try exact member code matching
                for country, code in COUNTRY_CODES.items():
                    if code in STANDARD_COLUMN_ORDER and member_code.upper() == country.replace(' ', '').upper()[:3]:
                        # Direct 3-letter code match (AUS, AUT, BEL, etc.)
                        if member_code.upper() in ['AUS', 'AUT', 'BEL', 'BRA', 'CAN', 'CHL', 'CHN', 'CYP', 'DNK', 'FIN', 'FRA', 'DEU', 'GRC', 'IND', 'IRL', 'ISR', 'ITA', 'JPN', 'KOR', 'KWT', 'LUX', 'MYS', 'MEX', 'NLD', 'NZL', 'NOR', 'PHL', 'POL', 'PRT', 'RUS', 'SAU', 'SGP', 'ZAF', 'ESP', 'SWE', 'CHE', 'THA', 'GBR', 'USA']:
                            # Create mapping based on exact code
                            code_to_country_map = {
                                'AUS': 'Australia',
                                'AUT': 'Austria', 
                                'BEL': 'Belgium',
                                'BRA': 'Brazil',
                                'CAN': 'Canada',
                                'CHL': 'Chile',
                                'CHN': 'China',
                                'CYP': 'Cyprus',
                                'DNK': 'Denmark',
                                'FIN': 'Finland',
                                'FRA': 'France',
                                'DEU': 'Germany',
                                'GRC': 'Greece',
                                'IND': 'India',
                                'IRL': 'Ireland',
                                'ISR': 'Israel',
                                'ITA': 'Italy',
                                'JPN': 'Japan',
                                'KOR': 'Korea',
                                'KWT': 'Kuwait',
                                'LUX': 'Luxembourg',
                                'MYS': 'Malaysia',
                                'MEX': 'Mexico',
                                'NLD': 'Netherlands',
                                'NZL': 'New Zealand',
                                'NOR': 'Norway',
                                'PHL': 'Philippines',
                                'POL': 'Poland',
                                'PRT': 'Portugal',
                                'RUS': 'Russian Federation',
                                'SAU': 'Saudi Arabia',
                                'SGP': 'Singapore',
                                'ZAF': 'South Africa',
                                'ESP': 'Spain',
                                'SWE': 'Sweden',
                                'CHE': 'Switzerland',
                                'THA': 'Thailand',
                                'GBR': 'United Kingdom',
                                'USA': 'United States'
                            }
                            
                            matched_country = code_to_country_map.get(member_code.upper())
                            if matched_country and matched_country in COUNTRY_CODES:
                                country_code = COUNTRY_CODES[matched_country]
                                if country_code in STANDARD_COLUMN_ORDER:
                                    idx = STANDARD_COLUMN_ORDER.index(country_code)
                                    data_row[idx] = amount
                                    print(f"[OK] Mapped {matched_country} ({member_name}): {amount:,.0f}")
                                    matched = True
                                    break
                
                # If no exact code match, try specific name matching with strict rules
                if not matched:
                    # Special cases that need exact matching
                    exact_matches = {
                        'Australia': ['Australia'],
                        'Austria': ['Austria'],
                        'Belgium': ['Belgium'],
                        'Hong Kong': ['Hong Kong Monetary Authority', 'HKMA'],
                        'Chile': ['Chile (Banco Central de Chile)'],
                        'Denmark': ['Denmark (Danmarks Nationalbank)'],
                        'Germany': ['Germany (Deutsche Bundesbank)'],
                        'Israel': ['Israel (Bank of Israel)'],
                        'Netherlands': ['Netherlands, The'],
                        'Philippines': ['Philippines (Bangko Sentral ng Pilipinas)'],
                        'Poland': ['Poland, Republic of (National Bank of Poland)'],
                        'Portugal': ['Portugal (Banco de Portugal)'],
                        'Sweden': ['Sweden (Sveriges Riksbank)'],
                        'Switzerland': ['Switzerland (Swiss National Bank)']
                    }
                    
                    for country, variants in exact_matches.items():
                        if country in COUNTRY_CODES:
                            for variant in variants:
                                if variant.lower() == member_name.lower() or variant in member_name:
                                    code = COUNTRY_CODES[country]
                                    if code in STANDARD_COLUMN_ORDER:
                                        idx = STANDARD_COLUMN_ORDER.index(code)
                                        data_row[idx] = amount
                                        print(f"[OK] Mapped {country} ({member_name}): {amount:,.0f}")
                                        matched = True
                                        break
                            if matched:
                                break
                
                # If still no match, try simple country name matching (but be very careful)
                if not matched:
                    simple_country_names = [
                        'Brazil', 'Canada', 'China', 'Cyprus', 'Finland', 'France', 
                        'Greece', 'India', 'Ireland', 'Italy', 'Japan', 'Korea', 
                        'Kuwait', 'Luxembourg', 'Malaysia', 'Mexico', 'New Zealand', 
                        'Norway', 'Russian Federation', 'Saudi Arabia', 'Singapore', 
                        'South Africa', 'Spain', 'Thailand', 'United Kingdom', 'United States'
                    ]
                    
                    for country in simple_country_names:
                        if country in COUNTRY_CODES and country.lower() == member_name.lower():
                            code = COUNTRY_CODES[country]
                            if code in STANDARD_COLUMN_ORDER:
                                idx = STANDARD_COLUMN_ORDER.index(code)
                                data_row[idx] = amount
                                print(f"[OK] Mapped {country} ({member_name}): {amount:,.0f}")
                                matched = True
                                break
                
                if not matched:
                    print(f"[ERROR] Could not map: {member_name} ({member_code}): {amount}")

        # 5. Create the final DataFrame with proper structure
        final_data = [header_row_1, header_row_2, data_row]
        combined_df = pd.DataFrame(final_data)

        # 6. Save the output to an Excel file with proper number formatting
        output_dir_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), OUTPUT_DIR)
        os.makedirs(output_dir_path, exist_ok=True)
        output_file_path = os.path.join(output_dir_path, OUTPUT_FILE)
        
        # Create Excel writer with formatting options
        with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
            combined_df.to_excel(writer, index=False, header=False, sheet_name='Sheet1')
            
            # Get the workbook and worksheet objects
            workbook = writer.book
            worksheet = writer.sheets['Sheet1']
            
            # Define number format with commas
            number_format = workbook.add_format({'num_format': '#,##0'})
            
            # Apply number formatting to data row (row 3, columns B onwards)
            # Note: xlsxwriter uses 0-based indexing, so row 3 is index 2
            for col in range(1, len(STANDARD_COLUMN_ORDER)):  # Start from column 1 (B), skip date column
                # Convert string numbers back to numeric for proper Excel formatting
                if data_row[col] is not None:
                    try:
                        # Remove commas and convert to number for Excel
                        numeric_value = float(str(data_row[col]).replace(',', ''))
                        worksheet.write(2, col, numeric_value, number_format)
                    except:
                        # If conversion fails, write as text
                        worksheet.write(2, col, data_row[col])
            
            # Make sure date column stays as text
            worksheet.write(2, 0, report_date)

        print(f"\nSuccess! Mapped data has been saved to '{output_file_path}'")
        print("\nFile structure:")
        print("Row 1: Timeseries IDs (CODEs)")
        print("Row 2: Timeseries descriptions") 
        print(f"Row 3: Data values (with date {report_date} in first column)")
        
        # Print summary of data found (calculate from string amounts)
        non_empty_values = [val for val in data_row[1:] if val is not None]
        print(f"\nData summary:")
        print(f"- Report date: {report_date}")
        print(f"- Countries with data: {len(non_empty_values)}")
        if non_empty_values:
            # Convert comma-formatted strings back to numbers for total calculation
            total_amount = sum(float(str(val).replace(',', '')) for val in non_empty_values if val)
            print(f"- Total borrowing amount: {total_amount:,.0f}")

    except Exception as e:
        print(f"\nAn unexpected error occurred: {e}")
        import traceback
        traceback.print_exc()

if __name__ == '__main__':
    main()