import os
import glob
import pdfplumber
import pandas as pd
import re

def parse_line(line):
    """
    Standardizes the extraction of property labels, context, values, and units.
    Example: 'Conductor AC resistance at 50 Hz 5.61 Ohm/km' 
    -> ('Conductor AC resistance at 50 Hz (Ohm/km)', '5.61')
    """
    # Standardize common context terms to ensure consistent headers
    line = line.replace('\n', ' ').strip()
    line = re.sub(r'50\s*Hz', '50 Hz', line, flags=re.IGNORECASE)
    line = re.sub(r'20\s*°C', '20°C', line, flags=re.IGNORECASE)
    
    if not line:
        return None, None
        
    # Standardize units to extract from values and put into headers
    # Removed 'Hz' from units as it is usually part of the context/header
    units = ["mm²", "mm", "kg/100m", "m", "Ohm/km", "MOhm.km", "kV", "V", "kN", "kg","µF / km"]
    units_pattern = r'(' + '|'.join([re.escape(u) for u in units]) + r')'
    
    # Non-numeric keywords that are values
    keywords = [
        'Orange', 'Copper', 'Circular', 'V-90', 'PVC', "X-90",'Class 2', 'Yes', 'No', 
        'Rigid', 'Stranded', 'Sector', 'Steel wires', 'Steel wire', 'Steel', 'Armour', "Screen","Copper Tape"
    ]
    kw_pattern = r'(' + '|'.join(keywords) + r')'

    # Case 1: Number + Unit (usually at the end)
    matches = list(re.finditer(r'([\d,./]+)\s*' + units_pattern + r'\b', line, re.IGNORECASE))
    if matches:
        m = matches[-1] # take the last one as the primary value
        val = m.group(1)
        unit = m.group(2)
        # Header is everything else combined
        prefix = line[:m.start()].strip()
        suffix = line[m.end():].strip()
        header = f"{prefix} {suffix}".strip()
        header = f"{header} ({unit})"
        return header.strip(), val

    # Case 2: Keywords at the end or following a space
    matches = list(re.finditer(r'\s' + kw_pattern + r'.*$', line, re.IGNORECASE))
    if not matches:
        matches = list(re.finditer(r'^' + kw_pattern + r'.*$', line, re.IGNORECASE))
    
    if matches:
        m = matches[-1]
        val = m.group(0).strip()
        header = line[:m.start()].strip()
        if not header:
            return line, ""
        return header, val

    # Case 3: Number only at the end
    matches = list(re.finditer(r'\s([\d,./]+)$', line))
    if matches:
        m = matches[-1]
        val = m.group(1)
        header = line[:m.start()].strip()
        return header, val

    return line, ""

def extract_lv_cables_data(folder_path, output_csv_path):
    all_rows = []

    pdf_files = glob.glob(os.path.join(folder_path, "*.pdf"))
    if not pdf_files:
        print(f"No PDF files found in {folder_path}")
        return

    for pdf_path in pdf_files:
        file_name = os.path.basename(pdf_path)
        print(f"Processing {file_name}...")

        try:
            with pdfplumber.open(pdf_path) as pdf:
                if len(pdf.pages) < 2:
                    continue

                page = pdf.pages[1]
                tables = page.extract_tables()
                
                if not tables:
                    print(f"  No tables found in {file_name}")
                    continue

                row_data = {"File Name": file_name}
                
                for table in tables:
                    for row in table:
                        if not row or not any(row):
                            continue
                        
                        content = " ".join([str(cell) for cell in row if cell]).strip()
                        if not content or content.upper() == "CHARACTERISTICS":
                            continue

                        key, val = parse_line(content)
                        if key and val:
                            row_data[key] = val

                all_rows.append(row_data)

        except Exception as e:
            print(f"  Error processing {file_name}: {e}")

    if all_rows:
        df = pd.DataFrame(all_rows)
        # Clean up: remove columns that only have category headers (empty values)
        df = df.dropna(axis=1, how='all')
        
        # Standardize headers: ensure one space before unit brackets if they were merged poorly
        df.columns = [re.sub(r'(\w)\(', r'\1 (', c) for c in df.columns]
        
        # Sort columns with File Name first
        cols = ['File Name'] + sorted([c for c in df.columns if c != 'File Name'])
        df = df[cols]
        
        df.to_csv(output_csv_path, index=False)
        print(f"\nSuccessfully created {output_csv_path} with {len(all_rows)} rows and {len(df.columns)} standardized columns.")
    else:
        print("\nNo data was extracted from the PDFs.")

if __name__ == "__main__":
    current_dir = os.path.dirname(os.path.abspath(__file__))
    lv_cables_folder = os.path.join(current_dir, "LV Cables")
    output_path = os.path.join(lv_cables_folder, "LV_Cables_Consolidated.csv")
    
    extract_lv_cables_data(lv_cables_folder, output_path)
