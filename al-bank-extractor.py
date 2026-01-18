import pdfplumber
import re
import argparse
import os
import sys
import pandas as pd

def parse_danish_number(num_str):
    """Converts '1.250,00' or '646,43-' to float."""
    if not num_str: return 0.0
    clean_str = num_str.strip().replace(' ', '')
    is_negative = clean_str.endswith('-')
    if is_negative:
        clean_str = clean_str[:-1]
    
    clean_str = clean_str.replace('.', '').replace(',', '.')
    
    try:
        val = float(clean_str)
        return -val if is_negative else val
    except ValueError:
        return 0.0

def clean_description(text):
    """Removes trailing value dates (e.g. ' 05.05') from descriptions."""
    if not text: return ""
    return re.sub(r'\s\d{2}\.\d{2}$', '', text).strip()

def extract_to_excel(pdf_path, output_xlsx, do_clean=False):
    print(f"--- Processing: {pdf_path} ---")
    print(f"--- Mode: {'Cleaning Descriptions' if do_clean else 'Raw Descriptions'} ---")
    
    transactions = []
    
    # Regex Pattern: Date ... Description ... Amount ... Balance
    line_pattern = re.compile(r"(\d{2}\.\d{2}\.\d{4})\s+(.+?)\s+([\d\.]+,\d{2}-?)\s+([\d\.]+,\d{2}-?)$")

    try:
        with pdfplumber.open(pdf_path) as pdf:
            for i, page in enumerate(pdf.pages):
                text = page.extract_text()
                if not text: continue
                
                lines = text.split('\n')
                for line in lines:
                    line = line.strip()
                    match = line_pattern.search(line)
                    
                    if match:
                        date_str, raw_desc, amount_str, balance_str = match.groups()
                        
                        final_desc = clean_description(raw_desc) if do_clean else raw_desc.strip()
                        
                        transactions.append({
                            'Date': date_str,
                            'Description': final_desc,
                            'Amount': parse_danish_number(amount_str),
                            'Balance': parse_danish_number(balance_str),
                        })
    except Exception as e:
        print(f"Error reading PDF file: {e}")
        sys.exit(1)

    if transactions:
        try:
            # Create DataFrame
            df = pd.DataFrame(transactions)
            
            # Convert Date string to real Datetime objects for Excel
            df['Date'] = pd.to_datetime(df['Date'], format='%d.%m.%Y')
            
            # Write to Excel
            df.to_excel(output_xlsx, index=False)
            
            print(f"Success: {len(transactions)} transactions extracted.")
            print(f"Saved to: {output_xlsx}")
            
        except ImportError:
            print("Error: Pandas or openpyxl not installed. Run: py -m pip install pandas openpyxl")
        except Exception as e:
            print(f"Error writing Excel file: {e}")
    else:
        print("No transactions found.")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Extract Danish bank statements to Excel.')
    parser.add_argument('-f', '--file', type=str, required=True, help='Path to PDF file')
    parser.add_argument('--clean', action='store_true', help='Clean trailing dates from descriptions')
    
    args = parser.parse_args()
    
    base_name = os.path.splitext(args.file)[0]
    output_file = f"{base_name}.xlsx"
    
    if os.path.exists(args.file):
        extract_to_excel(args.file, output_file, do_clean=args.clean)
    else:
        print(f"Error: File '{args.file}' not found.")