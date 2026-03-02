import pdfplumber
import pandas as pd
import os

def extract_tables_from_pdf(pdf_path, output_csv_path, start_page, end_page):
    all_data = []
    
    with pdfplumber.open(pdf_path) as pdf:
        # Pages are 0-indexed in pdfplumber, so page 6 is index 5
        # page 6 to 66 means index 5 to 65
        for i in range(start_page - 1, end_page):
            print(f"Extracting tables from page {i + 1}...")
            page = pdf.pages[i]
            tables = page.extract_tables()
            
            for table in tables:
                if table and len(table) > 1:
                    # Convert table to DataFrame
                    # Using the first row as columns, but handling potential None values
                    cols = [str(c).replace('\n', ' ').strip() if c is not None else f"Unnamed_{j}" for j, c in enumerate(table[0])]
                    df = pd.DataFrame(table[1:], columns=cols)
                    
                    # Log the columns to help debug if needed
                    # print(f"  Table columns: {cols}")
                    all_data.append(df)
    
    if all_data:
        # Combine all dataframes. Using join='outer' and sort=False to be safe
        # If headers are identical, they will align. If not, new columns will be created.
        final_df = pd.concat(all_data, ignore_index=True, sort=False)
        # Drop rows that are completely empty
        final_df = final_df.dropna(how='all')
        # Save to CSV
        final_df.to_csv(output_csv_path, index=False)
        print(f"Successfully extracted tables to {output_csv_path}")
    else:
        print("No tables found in the specified range.")

if __name__ == "__main__":
    current_dir = os.path.dirname(os.path.abspath(__file__))
    pdf_filename = "Medium Voltage_Guide_V1.pdf"
    pdf_path = os.path.join(current_dir, pdf_filename)
    
    output_dir = os.path.join(current_dir, "Extracted Files")
    output_filename = os.path.splitext(pdf_filename)[0] + ".csv"
    output_path = os.path.join(output_dir, output_filename)
    
    # Extract from pages 6 to 66
    extract_tables_from_pdf(pdf_path, output_path, 6, 66)
