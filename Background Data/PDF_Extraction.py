import pdfplumber
import pandas as pd
import os

def extract_tables_from_pdf(pdf_path, output_csv_path, start_page, end_page):
    all_rows = []
    
    with pdfplumber.open(pdf_path) as pdf:
        for i in range(start_page - 1, end_page):
            print(f"Extracting tables from page {i + 1}...")
            page = pdf.pages[i]
            # Use a slightly more robust table extraction strategy
            tables = page.extract_tables({
                "vertical_strategy": "lines",
                "horizontal_strategy": "lines",
                "intersection_y_tolerance": 10
            })
            
            # If no tables found with lines, try the default strategy
            if not tables:
                tables = page.extract_tables()
            
            for table in tables:
                if not table or len(table) < 2:
                    continue
                
                # 1. Identify the "Product Code" from Column C (index 2) of the first row
                # As per user request: "values currently shown in column C (first row of each table)"
                try:
                    product_code_from_c = str(table[0][2]).replace('\n', ' ').strip() if len(table[0]) > 2 and table[0][2] else ""
                except Exception:
                    product_code_from_c = ""

                # 2. Extract headers from the first column (Column 0)
                # These are the attributes like "Nominal Conductor Area (mm2)"
                headers = [str(row[0]).replace('\n', ' ').strip() if row[0] is not None else "" for row in table]
                
                # 3. Iterate through all data columns (starting from index 2, which is Column C)
                # We want to extract ALL variations (sizes) shown in the table
                num_cols = max(len(r) for r in table)
                for col_idx in range(2, num_cols):
                    # For each column, we create a new row in our final data
                    row_data = {
                        "Table Product Code": product_code_from_c,
                        "Page": i + 1
                    }
                    
                    has_actual_data = False
                    for row_idx, header in enumerate(headers):
                        if not header:
                            continue
                        
                        if row_idx < len(table) and col_idx < len(table[row_idx]):
                            val = table[row_idx][col_idx]
                            if val is not None and str(val).strip() != "":
                                # Check if the value is useful (not just a repeat of a header or "Unnamed")
                                val_str = str(val).replace('\n', ' ').strip()
                                if val_str and not val_str.startswith("Unnamed"):
                                    has_actual_data = True
                                row_data[header] = val_str
                            else:
                                row_data[header] = ""
                    
                    # Only add rows that have at least some data (e.g., an Area or Diameter)
                    if has_actual_data:
                        all_rows.append(row_data)
    
    if all_rows:
        final_df = pd.DataFrame(all_rows)
        
        # Clean up column names: remove duplicates if any, and handle 'Product Code' conflicts
        # We want the "Table Product Code" to be clear
        if "Product Code" in final_df.columns:
            # If there was a row named "Product Code", it's already in the data
            pass
        
        # Save to CSV
        final_df.to_csv(output_csv_path, index=False)
        print(f"Successfully extracted {len(final_df)} data rows to {output_csv_path}")
    else:
        print("No tables found in the specified range.")

if __name__ == "__main__":
    current_dir = os.path.dirname(os.path.abspath(__file__))
    pdf_filename = "Medium Voltage_Guide_V1.pdf"
    pdf_path = os.path.join(current_dir, pdf_filename)
    
    output_dir = os.path.join(current_dir, "Extracted Files")
    output_path = os.path.join(output_dir, pdf_filename.replace(".pdf", ".csv"))
    
    # Extract from pages 6 to 66
    extract_tables_from_pdf(pdf_path, output_path, 6, 66)
