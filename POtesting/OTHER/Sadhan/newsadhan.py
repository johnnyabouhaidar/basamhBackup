import pdfplumber

pdf_path = "2311002875094.pdf"
csv_path="output.xlsx"

with pdfplumber.open(pdf_path) as pdf:
    # Iterate through each page
    for page in pdf.pages:
        # Extract tables from the page
        tables = page.extract_tables(use_line_edges=True)
        
        # Iterate through each table
        for table in tables:
            
            for row in table:
                print(row)
                
            print("-" * 20)  # Separate tables with dashes

