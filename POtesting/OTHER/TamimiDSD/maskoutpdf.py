import fitz

def mask_region_in_pdf(input_path, output_path, x1, y1, x2, y2):
    pdf_document = fitz.open(input_path)

    for page_num in range(pdf_document.page_count):
        page = pdf_document.load_page(page_num)
        
        # Define a rectangle based on coordinates
        rect = fitz.Rect(x1, y1, x2, y2)
        
        # Create a redaction annotation on the specified region
        redact_annot = page.add_redact_annot(quad=rect,fill=(255,255,255))
        redact_annot.update()
        page.apply_redactions()
        page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE)
        

    pdf_document.save(output_path)
    pdf_document.close()

# Usage example
input_file_path = 'TM160-1015649.pdf'
output_file_path = 'masked_output.pdf'
x1, y1, x2, y2 = 400, 300, 140, 600  # Replace with your coordinates

mask_region_in_pdf(input_file_path, output_file_path, x1, y1, x2, y2)